using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
//using System.Threading.Tasks;

namespace PrinterLib
{
    public class PrintLPT : PrintBase
    {
        private string port = "lpt1";
        private int iHandle;

        public PrintLPT(string port)
        {
            this.port = port;
        }

        public bool Open()
        {
            iHandle = CreateFile(this.port, 0x40000000, 0, 0, 3, 0, 0);
            if (iHandle != -1)
                return true;
            else
                return false;
        }

        public bool Close()
        {
            return CloseHandle(iHandle);
        }

        public override bool Print()
        {
            return PrintContent(Encoding.Default);
        }

        public override bool Print(Encoding encode)
        {
            return PrintContent(encode);
        }

        public bool PrintContent(Encoding encode)
        {
            int i = 0;
            this.message = string.Empty;

            if (Open() == false)
            {
                message = port + " can not open";
                return false;
            }

            OVERLAPPED x = new OVERLAPPED();
            byte[] bytes = encode.GetBytes(content);
            for (int iQty = 0; iQty < labelQty; iQty++)
            {
                if (WriteFile(iHandle, bytes, bytes.Length, ref i, ref x) == false)
                {
                    message = port + " out put printer error!";
                    return false;
                }

                Thread.Sleep(30);
            }

            if (Close() == false)
            {
                message = port + "can not close";
                return false;
            }

            return true;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct OVERLAPPED
        {
            int Internal;
            int InternalHigh;
            int Offset;
            int OffSetHigh;
            int hEvent;
        }

        [DllImport("kernel32.dll")]
        private static extern int CreateFile(string lpFileName,
                             uint dwDesiredAccess,
                             int dwShareMode,
                             int lpSecurityAttributes,
                             int dwCreationDisposition,
                             int dwFlagsAndAttributes,
                             int hTemplateFile
                            );

        [DllImport("kernel32.dll")]
        private static extern bool WriteFile(int hFile,
                                     byte[] lpBuffer,
                                     int nNumberOfBytesToWrite,
                                     ref   int lpNumberOfBytesWritten,
                                     ref   OVERLAPPED lpOverlapped
                                );

        [DllImport("kernel32.dll")]
        private static extern bool CloseHandle(int hObject);
    }
}
