using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

// Serves only the owned static OBJECT page. Office files remain in the host.
internal sealed class LoopbackPage : IDisposable
{
    private readonly TcpListener listener = new TcpListener(IPAddress.Loopback, 0);
    private readonly string requestPath = "/" + Guid.NewGuid().ToString("N") + "/";
    private readonly byte[] body;
    private readonly string authority;
    private volatile bool disposed;
    private int served;
    internal readonly Uri Url;
    internal bool Served { get { return Volatile.Read(ref served) > 0; } }

    internal LoopbackPage(string html)
    {
        body = Encoding.UTF8.GetBytes(html);
        listener.Server.ExclusiveAddressUse = true;
        listener.Start(4);
        authority = "localhost:" + ((IPEndPoint)listener.LocalEndpoint).Port;
        Url = new Uri("http://" + authority + requestPath);
        Task.Run((Action)Serve);
    }

    private void Serve()
    {
        // The listener reserves the port for this document until the host closes.
        while (!disposed)
        {
            try
            {
                using (var client = listener.AcceptTcpClient())
                using (var stream = client.GetStream())
                {
                    stream.ReadTimeout = 1000;
                    stream.WriteTimeout = 1000;
                    var headers = new StringBuilder();
                    for (int n = 0; n < 4096; n++)
                    {
                        int value = stream.ReadByte();
                        if (value < 0) break;
                        headers.Append((char)value);
                        int length = headers.Length;
                        if (length >= 4 && headers[length - 4] == '\r' && headers[length - 3] == '\n' &&
                            headers[length - 2] == '\r' && headers[length - 1] == '\n') break;
                    }
                    string[] lines = headers.ToString().Split(new[] { "\r\n" }, StringSplitOptions.None);
                    bool expectedHost = Array.Exists(lines, line => String.Equals(line, "Host: " + authority, StringComparison.OrdinalIgnoreCase));
                    bool correctPage = expectedHost && (lines[0] == "GET " + requestPath + " HTTP/1.1" || lines[0] == "GET " + requestPath + " HTTP/1.0");
                    byte[] response = Encoding.ASCII.GetBytes(correctPage ? "HTTP/1.1 200 OK\r\nConnection: close\r\nCache-Control: no-store\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: " + body.Length + "\r\n\r\n" :
                        "HTTP/1.1 404 Not Found\r\nConnection: close\r\nContent-Length: 0\r\n\r\n");
                    stream.Write(response, 0, response.Length);
                    if (correctPage) { stream.Write(body, 0, body.Length); Interlocked.Increment(ref served); }
                }
            }
            catch (SocketException) { if (disposed) return; }
            catch (ObjectDisposedException) { return; }
            catch (IOException) { if (disposed) return; }
        }
    }

    public void Dispose()
    {
        disposed = true;
        listener.Stop();
    }
}
