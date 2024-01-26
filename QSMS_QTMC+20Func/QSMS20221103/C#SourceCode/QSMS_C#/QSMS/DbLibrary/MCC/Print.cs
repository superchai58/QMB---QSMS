using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;

namespace QSMS.DbLibrary.MCC
{
    public class Print
    {
        public class CommControl
        {

            private System.IO.Ports.SerialPort serialPort;
            //public CommControl(string PortName)
            public CommControl(string PortName, int BaudRate)
            {
                serialPort = new System.IO.Ports.SerialPort(PortName, BaudRate);

                //serialPort.BaudRate = PortName;
                //serialPort.DataBits = BaudRate;
                serialPort.StopBits = StopBits.One;
                serialPort.Parity = Parity.None;
                serialPort.Open();
            }

            public bool Dispose()
            {
                bool Result = true;
                if (serialPort != null)
                {
                    if (serialPort.IsOpen)
                    {
                        try
                        {
                            serialPort.Close();
                            serialPort.Dispose();
                        }
                        catch { Result = false; }
                    }
                }
                return Result;
            }

            public bool Write(string data)
            {
                try
                {
                    if (serialPort.IsOpen)
                    {
                        byte[] bData = System.Text.Encoding.Default.GetBytes(data);
                        serialPort.Write(bData, 0, bData.Length);
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch { return false; }
            }


        }
    }
}
