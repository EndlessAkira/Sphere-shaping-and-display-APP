using System.Net.Sockets;
using System.Net;
using System.Text;

namespace Server
{
    internal class Server
    {
        // Порт для прослушивания подключений
        private const int PORT = 8888;
        // Длинна очереди
        private const int _lengthQueue = 10;
        // Размер сообщения
        private const int SIZE = 1024;

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Запуск сервера");
                // Создание конечной точки по IP и порту
                IPEndPoint _iPEndPoint = new IPEndPoint(IPAddress.Any, PORT);
                // Создание сокета(v4, потоковый, TCP)
                Socket _socket1 = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                // Связь с конечной локальной точкой для ожидания вход. запросов
                _socket1.Bind(_iPEndPoint);
                // Включение прослушивания
                _socket1.Listen(_lengthQueue);
                // Вывод инфо о сервере
                Console.WriteLine("Дата и время: " + DateTime.Now);
                Console.WriteLine("Прослушивающий сокет:\n      Дескриптор: {0}\n      IPv4: {1}\n      Порт: {2}", _socket1.Handle, _iPEndPoint.Address, _iPEndPoint.Port);

                Console.WriteLine("Сервер в режиме ожидания");
                // Инициализация клиентского сокета в случае подключения клиента к серверу
                Socket _socket2 = _socket1.Accept();
                // Инициализация переменной для сообщения от сервера
                String _dataRec = "";
                while (true)
                {
                    Console.WriteLine("Дата и время: " + DateTime.Now);
                    Console.WriteLine("Получение запроса от клиента:\n      Дескриптор: {0}\n      IPv4: {1}\n      Порт: {2}", _socket2.Handle, ((IPEndPoint)_socket2.RemoteEndPoint).Address, ((IPEndPoint)_socket2.RemoteEndPoint).Port);
                    byte[] _byteRec = new byte[SIZE];
                    // Приём сообщения от клиента, запись сообщения и его длинны
                    int _lenBytesReciver = _socket2.Receive(_byteRec);
                    // Декодировка
                    _dataRec += Encoding.ASCII.GetString(_byteRec, 0, _lenBytesReciver);
                    // Есть ли в сообщении еще символы
                    if (_dataRec.IndexOf('.') > -1)
                    {
                        break;
                    }
                }
                Console.WriteLine("Получено сообщение от клиента: {0}", _dataRec);
                // Инициализация сообщения для клиента
                string dataSend = $"1|15|25|25|2|90|0|90|Red";
                
                // Кодировка сообщения
                byte[] byteSend = Encoding.ASCII.GetBytes(dataSend);
                // Отправка сообщения клиенту
                int lenBytesSend = _socket2.Send(byteSend);
                Console.WriteLine("Отправка клиенту {0} bytes", lenBytesSend);
                // Инициирование закрытия сокета клиента
                _socket2.Shutdown(SocketShutdown.Both);
                // Закрытие сокета клиента
                _socket2.Close();
                Console.WriteLine("Общение с клиентом остановлено");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                Console.ReadLine();
            }
        }
    }
}

