using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
//using System.Threading.Tasks;

namespace PrinterLib
{
    public class PrintCom : PrintBase
    {
        private SerialPort serialPort = null;

        public PrintCom(SerialPort serialPort)
        {
            this.serialPort = serialPort;
        }

        public PrintCom(string port, string commSetting)
        {
            this.serialPort = GenSerialPort(port, commSetting);
        }

        public override bool Print()
        {
            return PrintContent(Encoding.Default);
        }

        public override bool Print(Encoding encode)
        {
            return PrintContent(encode);
        }

        private bool PrintContent(Encoding encode)
        {
            if (serialPort == null)
            {
                message = "SerialPort 配置不正确,请确认";
                return false;
            }

            if (serialPort.IsOpen == false)
            {
                try
                {
                    serialPort.Open();
                }
                catch (Exception ex)
                {
                    message = "SerialPort打开失败 " + ex.Message;
                    return false;
                }
            }
            byte[] bytes = System.Text.Encoding.Default.GetBytes(content);
            for (int iQty = 0; iQty < labelQty; iQty++)
            {
                serialPort.Write(bytes, 0, bytes.Length);
                Thread.Sleep(30);
            }
            serialPort.Close();
            return true;
        }

        private SerialPort GenSerialPort(string port, string commonSetting)
        {
            Parity parity;
            StopBits stopBits;

            string[] arraySettings = commonSetting.Split(',');
            if (arraySettings.Length == 4)
            {
                switch (arraySettings[1].ToUpper())
                {
                    case "N":
                        parity = Parity.None;
                        break;
                    case "E":
                        parity = Parity.Even;
                        break;
                    case "M":
                        parity = Parity.Mark;
                        break;
                    case "O":
                        parity = Parity.Odd;
                        break;
                    case "S":
                        parity = Parity.Space;
                        break;
                    default:
                        parity = Parity.None;
                        break;
                }

                switch (arraySettings[3].ToUpper())
                {
                    case "1":
                        stopBits = StopBits.One;
                        break;
                    case "2":
                        stopBits = StopBits.Two;
                        break;
                    default:
                        stopBits = StopBits.None;
                        break;
                }

                return new SerialPort(
                    port, int.Parse(arraySettings[0]), parity, int.Parse(arraySettings[2]), stopBits);
            }

            return null;
        }
    }
}
