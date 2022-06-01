using System;
using System.Net;
using System.Net.Mail;
using System.Configuration;
using System.IO;

namespace General.Librerias.CodigoUsuario
{
    public class Correo
    {
        public static string EnviarMensaje(beMensaje obeMensaje)
        {
            string rpta = "";
            MailMessage msg = null;
            try
            {
                string servidorCorreoIp = ConfigurationManager.AppSettings["ServidorCorreoIp"];
                string servidorCorreoPuerto = ConfigurationManager.AppSettings["ServidorCorreoPuerto"];
                string correoUsuario = ConfigurationManager.AppSettings["CorreoUsuario"];
                string archivoClave = ConfigurationManager.AppSettings["ArchivoPwd"];
                string correoClave = File.ReadAllText(archivoClave);
                SmtpClient client = new SmtpClient();
                client.Credentials = new NetworkCredential(correoUsuario, correoClave);
                //client.Timeout = 240;
                client.Port = int.Parse(servidorCorreoPuerto);
                client.Host = servidorCorreoIp;
                client.EnableSsl = obeMensaje.EsSeguro;
                msg = new MailMessage();
                msg.From = new MailAddress(obeMensaje.De);
                foreach (string para in obeMensaje.Para)
                {
                    msg.To.Add(new MailAddress(para));
                }
                foreach (string archivo in obeMensaje.Archivo)
                {
                    msg.Attachments.Add(new Attachment(archivo));
                }
                msg.Subject = obeMensaje.Asunto;
                msg.IsBodyHtml = true;
                msg.Body = obeMensaje.Contenido;
                client.Send(msg);
                rpta = "Se envio el Correo satisfactoriamente";
            }
            catch (Exception ex)
            {
                rpta = "Error al Enviar Correo: " + ex.Message;
            }
            return rpta;
        }
    }
}
