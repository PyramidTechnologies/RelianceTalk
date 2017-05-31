#region Copyright & License
/*
MIT License

Copyright (c) 2017 Pyramid Technologies

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
 */
#endregion

using System;
using System.IO;
using System.IO.Ports;
using System.Runtime.InteropServices;
using System.Threading;
using ThermalTalk;
namespace ThermalConsole
{
    class Program
    {
        private static SerialPort _mPort;

        static void Main(string[] args)
        {
            if (args == null || args.Length < 3)
            {
                PrintBadArgs();
                return;
            }

            // Get args
            var file = args[0];
            var delay = args[1];
            var uuid = args[2];
            var repeat = args.Length > 3 ? args[3] : "-1";

            // Read data
            byte[] raw = null;
            try
            {
                raw = File.ReadAllBytes(file);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error reading receipt template: {0}", e.Message);
                return;
            }

            Opts opts = null;
            try
            {

                opts = new Opts()
                {
                    FileName = file,
                    Data = raw,
                    DeviceId = uuid,
                    SecondsDelay = Int32.Parse(delay),
                    RepeatCount = Int32.Parse(repeat),                           
                    CounterIndex = 0,
                };

                // Find where the counter should go. 3 in row of 0xBB
                int index = 0;
                foreach(var b in raw)
                {
                    index++;
                    if (index < 3) continue;
                    if (b != 0xBB) continue;
                    if (raw[index - 2] != 0xBB && raw[index - 1] != 0xBB) continue;
                    opts.CounterIndex = index;
                    break;
                }

                if(opts.SecondsDelay < 7)
                {
                    Console.WriteLine("Minimum delay is 7 seconds, setting to minimum");
                    opts.SecondsDelay = 7;
                }

            }
            catch (Exception e)
            {
                Console.WriteLine("Invalid settings: {0}", e.Message);
                PrintBadArgs();
                return;
            }

            if (uuid.StartsWith("COM", StringComparison.CurrentCultureIgnoreCase))
            {
                SerialPrintMode(opts);
            }
            else
            {
                WindowsPrintMode(opts);
            }
        }

        private static void PrintBadArgs()
        {
            Console.WriteLine("ThermalConsole  © Pyramid Technologies Inc. 2017");
            Console.WriteLine("Usage: ");
            Console.WriteLine("\t path/to/test/file <send_delay> <device_id> <optional_repeat_count>\n");
            Console.WriteLine("\ttest file is the ESC/POS receipt to print");
            Console.WriteLine("\tsend_delay is the time in seconds before repeating print job");
            Console.WriteLine("\tdevice_id Windows printer name or serial port name. Serial port name must start with COM");
            Console.WriteLine("optional_repeat_count is the optional number of times to print the receipt. omit to print forever.");
        }

        private static void SerialPrintMode(Opts opts)
        {
            // TODO make baud rate configurable
            _mPort = new SerialPort(opts.DeviceId, 19200);
            _mPort.DataBits = 8;
            _mPort.Handshake = Handshake.None;
            _mPort.WriteTimeout = 1000;
            _mPort.ReadTimeout = 1000;
            _mPort.WriteBufferSize = 4 * 1024;
            _mPort.ReadBufferSize = 4 * 1024;
            _mPort.Encoding = System.Text.Encoding.GetEncoding("Windows-1252");
            _mPort.DtrEnable = true;
            _mPort.RtsEnable = true;
            _mPort.DiscardNull = false;

            var data = new byte[opts.Data.Length];
            Array.Copy(opts.Data, data, data.Length);

            try
            {
                _mPort.Open();

                var c = 0;
                while(opts.RepeatCount == -1 || --opts.RepeatCount > 0)
                {
                    // No counter to inject
                    if (opts.CounterIndex == 0)
                    {
                        _mPort.Write(data, 0, data.Length);
                    }
                    else
                    {
                        // Inject a counter
                        _mPort.Write(data, 0, opts.CounterIndex-2);
                        _mPort.Write(string.Format(" {0}", (++c).ToString("D8")));
                        _mPort.Write(data, opts.CounterIndex+2, data.Length - opts.CounterIndex - 2);
                    }

                    Thread.Sleep(opts.SecondsDelay * 1000);
                }                

            } 
            catch(Exception e)
            {
                Console.WriteLine("Failed to write data: {0}", e.Message);
                Console.WriteLine("Opts: {0}", opts.ToString());
            }
        }

        private static void WindowsPrintMode(Opts opts)
        {
            // Gotta get a pointer on the local heap. Fun fact, the naming suggests that
            // this would be on the stack but it isn't. Windows no longer has a global heap
            // per se so these naming conventions are legacy cruft.
            IntPtr ptr = Marshal.AllocHGlobal(opts.Data.Length);
            Marshal.Copy(opts.Data, 0, ptr, opts.Data.Length);

            try
            {
                while (opts.RepeatCount == -1 || --opts.RepeatCount > 0)
                {
                    bool result = RawPrinterHelper.SendBytesToPrinter(opts.DeviceId, ptr, opts.Data.Length);
                    Thread.Sleep(opts.SecondsDelay * 1000);
                }          
            }
            catch (Exception e)
            {
                Console.WriteLine("Error writing to Windows Printer: {0}", e.Message);
            }
            finally
            {
                Marshal.FreeHGlobal(ptr);
            }
        }
    }

    class Opts
    {
        public string FileName { get; set; }
        public byte[] Data { get; set; }
        public string DeviceId { get; set; }
        public int SecondsDelay { get; set; }
        public int RepeatCount { get; set; }
        public int CounterIndex { get; set; }

        public override string ToString()
        {
            return string.Format("File: {0}, DeviceId: {1}, SecondsDelay{2}, RepeatCount{3}",
                FileName, DeviceId, SecondsDelay, RepeatCount);
        }
    }
}
