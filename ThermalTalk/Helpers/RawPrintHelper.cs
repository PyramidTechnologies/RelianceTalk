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
namespace ThermalTalk
{
    using System;
    using System.IO;
    using System.Runtime.InteropServices;

    /// <summary>
    /// PInvoke wrapper for winspool driver functions
    /// </summary>
    public class RawPrinterHelper
    {
        // no ctor
        private RawPrinterHelper() { }
     
        /// <summary>
        /// Send unmanaged data to the target printer.
        /// When the function is given a printer name and an unmanaged array
        /// of bytes, the function sends those bytes to the print queue.
        /// Returns true on success, false on failure.
        /// </summary>
        /// <param name="szPrinterName">String name of printer</param>
        /// <param name="pBytes">Pointer to data</param>
        /// <param name="dwCount">Length of data in bytes</param>
        /// <returns>bool</returns>
        public static bool SendBytesToPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            Int32 dwError = 0, dwWritten = 0;
            IntPtr hPrinter = new IntPtr(0);
            NativeMethods.DOCINFOA di = new NativeMethods.DOCINFOA();

            bool bSuccess = false; // Assume failure unless you specifically succeed.

            di.pDocName = "ESCPOSTTester";
            di.pDataType = "RAW";
            var normalized = szPrinterName.Normalize();

            // Open the printer.
            if (NativeMethods.OpenPrinter(normalized, out hPrinter, IntPtr.Zero))
            {
                // Start a document.
                if (NativeMethods.StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (NativeMethods.StartPagePrinter(hPrinter))
                    {
                        bSuccess = NativeMethods.WritePrinter(hPrinter, pBytes, dwCount, out dwWritten);
                        NativeMethods.EndPagePrinter(hPrinter);
                    }
                    NativeMethods.EndDocPrinter(hPrinter);
                }
                NativeMethods.ClosePrinter(hPrinter);
            }

            // If you did not succeed, GetLastError may give more information
            // about why not.
            if (bSuccess == false)
            {
                dwError = Marshal.GetLastWin32Error();
            }
            return bSuccess;
        }

        /// <summary>
        /// Send file to printer as a print job. If the path does not exists
        /// an exception will be thrown
        /// </summary>
        /// <param name="szPrinterName">String name of printer</param>
        /// <param name="szFileName">Path to file to print</param>
        /// <returns>bool</returns>
        public static bool SendFileToPrinter(string szPrinterName, string szFileName)
        {
            // Open the file
            // MS for whatever reason, make the reader take ownership of this stream
            // so DO NOT use the using pattern here. This could lead to an ObjectDisposed
            // exception. Let the BinaryReader (br) clean up fs for us.
            FileStream fs = new FileStream(szFileName, FileMode.Open);
            using (BinaryReader br = new BinaryReader(fs))
            {
                // Dim an array of bytes big enough to hold the file's contents.
                Byte[] bytes = new Byte[fs.Length];
                bool bSuccess = false;
                // Your unmanaged pointer.
                IntPtr pUnmanagedBytes = new IntPtr(0);
                int nLength;

                nLength = Convert.ToInt32(fs.Length);

                // Read the contents of the file into the array.
                bytes = br.ReadBytes(nLength);

                // Allocate some unmanaged memory for those bytes.
                pUnmanagedBytes = Marshal.AllocCoTaskMem(nLength);

                // Copy the managed byte array into the unmanaged array.
                Marshal.Copy(bytes, 0, pUnmanagedBytes, nLength);

                // Send the unmanaged bytes to the printer.
                bSuccess = SendBytesToPrinter(szPrinterName, pUnmanagedBytes, nLength);

                // Free the unmanaged memory that you allocated earlier.
                Marshal.FreeCoTaskMem(pUnmanagedBytes);

                return bSuccess;
            }
        }

        /// <summary>
        /// Convenience wrapper around SendBytesToPrinter 
        /// </summary>
        /// <param name="szPrinterName">String name of printer</param>
        /// <param name="szString">String to send to printer</param>
        /// <returns></returns>
        public static bool SendStringToPrinter(string szPrinterName, string szString)
        {
            IntPtr pBytes;
            Int32 dwCount;
            // How many characters are in the string?
            dwCount = szString.Length;
            // Assume that the printer is expecting ANSI text, and then convert
            // the string to ANSI text.
            pBytes = Marshal.StringToCoTaskMemAnsi(szString);
            // Send the converted ANSI string to the printer.
            SendBytesToPrinter(szPrinterName, pBytes, dwCount);
            Marshal.FreeCoTaskMem(pBytes);
            return true;
        }

        public static bool ReadFromPrinter(string szPrinterName, IntPtr pBytes, Int32 dwCount)
        {
            // Read the data from the printer.
            Int32 dwError = 0, dwBytesRead = 0;
            IntPtr hPrinter = new IntPtr(0);
            NativeMethods.DOCINFOA di = new NativeMethods.DOCINFOA();
            NativeMethods.PRINTER_DEFAULTS pd = new NativeMethods.PRINTER_DEFAULTS();
            pd.DesiredAccess = 0x00000020;
         
            bool bSuccess = false;

            // Open the printer.
            if (NativeMethods.OpenPrinter2(szPrinterName.Normalize(), out hPrinter, pd))
            {
                // Start a document.
                if (NativeMethods.StartDocPrinter(hPrinter, 1, di))
                {
                    // Start a page.
                    if (NativeMethods.StartPagePrinter(hPrinter))
                    {
                        // read your bytes.
                        bSuccess = NativeMethods.ReadPrinter(hPrinter, out pBytes, dwCount, out dwBytesRead);
                        
                        // If you did not succeed, GetLastError may give more information
                        // about why not.
                        if (bSuccess == false)
                        {
                            dwError = Marshal.GetLastWin32Error();
                        }

                        NativeMethods.EndPagePrinter(hPrinter);
                    }
                    NativeMethods.EndDocPrinter(hPrinter);
                }
                NativeMethods.ClosePrinter(hPrinter);
            }

            return bSuccess;
        }
    }
}
