using System;
using System.Configuration;
using System.IO;

namespace General.Librerias.CodigoUsuario
{
    public class Log
    {
        public static void Grabar(beLog obeLog, bool cifrado = false)
        {
            try
            {
                string rutaLog = ConfigurationManager.AppSettings["RutaLog"];
                if(cifrado)
                {
                    rutaLog = Seguridad.descifrarAESHex(rutaLog);
                }
                DateTime fechaActual = DateTime.Now;
                string nombre = String.Format("Log_{0}_{1}_{2}.txt", fechaActual.Year, fechaActual.Month.ToString().PadLeft(2, '0'), fechaActual.Day.ToString().PadLeft(2, '0'));
                string archivoLog = Path.Combine(rutaLog, nombre);
                Objeto.Grabar(obeLog, archivoLog);
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }

        }
    }
}
