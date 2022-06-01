using System;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace General.Librerias.CodigoUsuario
{
    public class Compresion
    {
        public static byte[] ComprimirZip(string data)
        {
            byte[] rpta = null;
            try
            {
                byte[] buffer = Encoding.Default.GetBytes(data);
                using (MemoryStream msEntrada = new MemoryStream(buffer))
                {
                    using (MemoryStream msSalida = new MemoryStream())
                    {
                        using (GZipStream gzip = new GZipStream(msSalida, CompressionMode.Compress))
                        {
                            msEntrada.CopyTo(gzip);
                            rpta = msSalida.ToArray();
                            //rpta = Encoding.Default.GetString(buffer);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                beLog obeLog = new beLog();
                obeLog.MensajeError = ex.Message;
                obeLog.DetalleError = ex.StackTrace;
                Log.Grabar(obeLog);
                rpta = Encoding.Default.GetBytes("Error - No se pudo comprimir con Gzip");
            }
            return rpta;
        }

        public static string DescomprimirZip(string data)
        {
            string rpta = "";
            try
            {
                byte[] buffer = Encoding.Default.GetBytes(data);
                using (MemoryStream msEntrada = new MemoryStream(buffer))
                {
                    using (MemoryStream msSalida = new MemoryStream())
                    {
                        using (GZipStream gzip = new GZipStream(msEntrada, CompressionMode.Decompress))
                        {
                            gzip.CopyTo(msSalida);
                            buffer = msSalida.ToArray();
                            rpta = Encoding.Default.GetString(buffer);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                beLog obeLog = new beLog();
                obeLog.MensajeError = ex.Message;
                obeLog.DetalleError = ex.StackTrace;
                Log.Grabar(obeLog);
                rpta = "Error - No se pudo descomprimir con Gzip";
            }
            return rpta;
        }

        public static byte[] ComprimirDeflate(string data)
        {
            byte[] rpta = null;
            try
            {
                byte[] buffer = Encoding.Default.GetBytes(data);
                using (MemoryStream msEntrada = new MemoryStream(buffer))
                {
                    using (MemoryStream msSalida = new MemoryStream())
                    {
                        using (DeflateStream plain = new DeflateStream(msSalida, CompressionMode.Compress))
                        {
                            msEntrada.CopyTo(plain);
                            rpta = msSalida.ToArray();
                            //rpta = Encoding.Default.GetString(buffer);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                beLog obeLog = new beLog();
                obeLog.MensajeError = ex.Message;
                obeLog.DetalleError = ex.StackTrace;
                Log.Grabar(obeLog);
                rpta = Encoding.Default.GetBytes("Error - No se pudo comprimir con Deflate");
            }
            return rpta;
        }

        public static byte[] DescomprimirDeflate(byte[] buffer)
        {
            byte[] rpta = null;
            try
            {
                using (MemoryStream msEntrada = new MemoryStream(buffer))
                {
                    using (MemoryStream msSalida = new MemoryStream())
                    {
                        using (DeflateStream plain = new DeflateStream(msEntrada, CompressionMode.Decompress))
                        {
                            plain.CopyTo(msSalida);
                            rpta = msSalida.ToArray();
                            //rpta = Encoding.Default.GetString(buffer);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                beLog obeLog = new beLog();
                obeLog.MensajeError = ex.Message;
                obeLog.DetalleError = ex.StackTrace;
                Log.Grabar(obeLog);
                rpta = Encoding.Default.GetBytes("Error - No se pudo descomprimir con Deflate");
            }
            return rpta;
        }

        public static byte[] ComprimirDeflateBytes(byte[] buffer)
        {
            byte[] rpta = null;
            try
            {
                using (MemoryStream msEntrada = new MemoryStream(buffer))
                {
                    using (MemoryStream msSalida = new MemoryStream())
                    {
                        using (DeflateStream plain = new DeflateStream(msSalida, CompressionMode.Compress))
                        {
                            msEntrada.CopyTo(plain);
                            rpta = msSalida.ToArray();
                            //rpta = Encoding.Default.GetString(buffer);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                beLog obeLog = new beLog();
                obeLog.MensajeError = ex.Message;
                obeLog.DetalleError = ex.StackTrace;
                Log.Grabar(obeLog);
                rpta = Encoding.Default.GetBytes("Error - No se pudo comprimir con Deflate");
            }
            return rpta;
        }
    }
}
