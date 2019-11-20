using Newtonsoft.Json;
using Serilog;
using Signer.Impl;
using Signer.IsaHub.JsonClasses.Pdf;
using Signer.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reactive.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Websocket.Client;
using System.Configuration;
using System.Linq;

namespace Signer.IsaHub
{

    internal class IsaSigner : PdfSignerBase
    {
        ILogger logger = Log.ForContext<IsaSigner>();
        private static String uri = "wss://127.0.0.1";
        private static List<int> ports = new List<int>() { 4321, 4322, 4323 };

        private WebsocketClient client;

        public override string Pin { set => throw new NotImplementedException(); }

        static IsaSigner()
        {
            var configUri = ConfigurationManager.AppSettings.Get("IsaHubHostUri");
            var configPorts = ConfigurationManager.AppSettings.Get("IsaHubPorts");
            if (!String.IsNullOrWhiteSpace(configUri) && Uri.IsWellFormedUriString(configUri.Trim(), UriKind.Absolute))
                uri = configUri.Trim();

            List<int> portsFromConfig = new List<int>();

            if (!String.IsNullOrWhiteSpace(configPorts))
            {
                foreach (var sPort in configPorts.Split(','))
                    if (int.TryParse(sPort, out var iPort))
                        portsFromConfig.Add(iPort);
                if (portsFromConfig.Count > 0)
                    ports = portsFromConfig;
            }
        }

        public override event EventHandler<SignedFileEventArgs> SignedFileEvent;

        public IsaSigner()
        {
            //Intento de conexión al websocket iSCertHub - Se intenta con cada uno de los puertos configurados
            //hasta que la conexión sea exitosa o no haya mas puertos para probar, en cuyo caso devuelve una excepción.
            var intento = 1;
            var ok = false;
            var url = new Uri($"{uri}:{ports[0]}");
            
            while (!ok)
            {
                client = new WebsocketClient(url);
                try
                {
                    var cancellationTokenSource = new CancellationTokenSource(1500);
                    client.Start().Wait(cancellationTokenSource.Token);
                    ok = true;
                }

                catch (Exception exception)
                {
                    client.Dispose();
                    logger.Error(exception, "Error inesperado conectando al socket en la url {url}.", client.Url.ToString());
                    if (intento == ports.Count)
                    {
                        var errorMessage = "Error accediendo al servicio de firmas.\nVerifique que el isCertHub esté levantado y que la cédula está insertada en el lector";

                        throw new InstanceCreationException(errorMessage);
                    }
                    else
                    {
                        url = new Uri($"{uri}:{ports[intento]}");
                        intento++;
                    }
                }
            }

        }


        #region  Pdf



        public override async Task SignPdf(FileInfo fileInfo)
        {
            await SignPdf(fileInfo.FullName);
        }

        public override async Task SignPdf(string filePath)
        {

            await SignPdf(filePath, File.ReadAllBytes(filePath));
        }

        public override async Task SignPdf(string fileId, byte[] fileData)
        {

            var exitEvent = new ManualResetEvent(false);


            var suscr = client.MessageReceived
                 .Where(msg => msg.MessageType == System.Net.WebSockets.WebSocketMessageType.Text)
                 .Take(1)
                 .Subscribe(msg => ProcessPdfResponse(msg, fileId, exitEvent));

            await client.Send(GetPdfJsonParameter(fileData));

            //El websocket de isCertHub no está diseñado para recibir varios archivos. Se debe enviar uno y
            //aguardar la respuesta. Por ese motivo se debe bloquear la ejecución hasta recibir una respuesta,
            //que va a corresponder al archivo enviado. En caso que el socket no devuelva una respuesta esperada,
            //la aplicación queda bloqueada. No es válido poner un timeout y que se pueda continuar con el siguiente archivo,
            //pues la siguiente respuesta puede ser del archivo pendiente ya que no hay forma de identificar a que archivo 
            //pertenece el base64 del mensaje recibido. Mejorable.
            exitEvent.WaitOne();

        }

        private void ProcessPdfResponse(ResponseMessage msg, String fileId, ManualResetEvent exitEvent)
        {
            SignedFileEventArgs eventArgs;
            var res = JsonConvert.DeserializeObject<Result>(msg.ToString());
            if (res.code == 0)
            {
                eventArgs = new SignedFileEventArgs(res.FileData, fileId, false, res.message);
            }
            else
                eventArgs = new SignedFileEventArgs(null, fileId, true, res.message);
            OnRaiseSignedFileEvent(eventArgs);
            exitEvent.Set();
        }

        protected virtual void OnRaiseSignedFileEvent(SignedFileEventArgs e)
        {
            EventHandler<SignedFileEventArgs> handler = SignedFileEvent;

            if (handler != null)
            {
                handler(this, e);
            }
        }

        //Es una configuración por defecto para todos los pdf. Mas adelante se puede extender la clase y recibir mas información como parámetros y personalizarlo.
        private string GetPdfJsonParameter(byte[] fileData)
        {
            var config = new Config()
            {
                apariencia = false,
                modoFirma = 2,
                paginaFirma = -1,
                posicion = new List<int>() { 365, 50, 495, 80 },
                urlImagen = ""
            };
            var value = new Value()
            {
                config = config,
                content = Convert.ToBase64String(fileData)
            };

            var param = new Parameter()
            {
                module = "pdf-sign",
                value = JsonConvert.SerializeObject(value)
            };
            return JsonConvert.SerializeObject(param);
        }

        public override void Dispose()
        {
            Task.Delay(10).Wait();
            client.Dispose();
        }

       

        #endregion
    }
}

