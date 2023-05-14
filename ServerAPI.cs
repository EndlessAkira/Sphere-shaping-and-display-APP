using System;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace SphereBuilder
{
    abstract class ServerAPI
    {
        public static int[] GetMessageFromSocket(int port)
        {

            int[] arrError = { 228 };
            try
            {
                // Буфер для входящих данных
                byte[] bytes = new byte[1024];

                // Соединяемся с удаленным устройством

                // Устанавливаем удаленную точку для сокета
                IPHostEntry ipHost = Dns.GetHostEntry("localhost");
                IPAddress ipAddr = ipHost.AddressList[0];
                IPEndPoint ipEndPoint = new IPEndPoint(ipAddr, port);

                Socket sender = new Socket(ipAddr.AddressFamily, SocketType.Stream, ProtocolType.Tcp);

                // Соединяем сокет с удаленной точкой
                sender.Connect(ipEndPoint);


                // Получаем ответ от сервера
                int bytesRec = sender.Receive(bytes);


                string answer = Encoding.UTF8.GetString(bytes, 0, bytesRec);
                int[] arrayResult = answer.Split(' ').Select(int.Parse).ToArray();


                // Освобождаем сокет
                sender.Shutdown(SocketShutdown.Both);
                sender.Close();

                return arrayResult;
            }
            catch(Exception e)
            {
                return arrError;
            }
            
        }

        
    }
}
