using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
//using System.Threading.Tasks;

namespace PrinterLib
{
    public class PrintNetwork : PrintBase
    {
        private string ip = string.Empty;
        private int port = 80;

        public PrintNetwork(string ip)
        {
            this.ip = ip;
        }

        public PrintNetwork(string ip, string port)
        {
            this.ip = ip;
            this.port = int.Parse(port);
        }

        public override bool Print()
        {
            return PrintContent(Encoding.Default);
        }

        public override bool Print(Encoding encode)
        {
            return PrintContent(encode);
        }

        private bool PrintContent(Encoding encoder)
        {
            IPEndPoint hostEndPoint = new IPEndPoint(IPAddress.Parse(ip), port);
            Socket socket = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
            socket.Connect(hostEndPoint);

            if (socket.Connected == false)
            {
                message = string.Format("Can not connect Network printer,Ip:{0},Port:{1}", ip, port);
                return false;
            }
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            for (int i = 0; i < labelQty; i++)
            {
                socket.Send(bytes, bytes.Length, 0);
                Thread.Sleep(30);
            }

            if (socket.Connected == true)
            {
                socket.Shutdown(SocketShutdown.Both);
                socket.Close();
            }

            return true;
        }
    }
}
