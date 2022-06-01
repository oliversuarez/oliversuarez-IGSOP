using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;

namespace General.Librerias.CodigoUsuario
{
    public class Seguridad
    {
        public static string cifrarSHA256Hex(string mensaje)
        {
            string rpta = "";
            SHA256Managed sha = new SHA256Managed(); //32 bytes
            byte[] buffer = Encoding.Default.GetBytes(mensaje);
            byte[] hash = sha.ComputeHash(buffer);
            //Convertir cada byte en hexadecimal 32 x 2 = 64 bytes
            rpta = BitConverter.ToString(hash).Replace("-", "").ToLower();
            return rpta;
        }

        public static byte[] convertirHexadecimalToBytes(string mensajeHex)
        {
            int n = mensajeHex.Length;
            byte[] bytes = new byte[n / 2];
            for (int i = 0; i < n; i += 2)
                bytes[i / 2] = Convert.ToByte(mensajeHex.Substring(i, 2), 16);
            return bytes;
        }

        public static string cifrarAESHex(string mensaje)
        {
            string dataCifradaHex = "";
            AesCryptoServiceProvider aes = new AesCryptoServiceProvider();
            string strClave = "Curs0P@vn@vMJ".PadRight(32, ' ');
            string strVectorInicio = "Lduenas".PadRight(16, ' ');
            byte[] bufferClave = Encoding.Default.GetBytes(strClave);
            byte[] bufferVectorInicio = Encoding.Default.GetBytes(strVectorInicio);            
            using (MemoryStream ms = new MemoryStream())
            {
                try
                {
                    using (CryptoStream cs = new CryptoStream(ms, aes.CreateEncryptor(bufferClave, bufferVectorInicio), CryptoStreamMode.Write))
                    {

                        using (StreamWriter sw = new StreamWriter(cs))
                        {
                            sw.Write(mensaje);
                        }
                        byte[] bufferCifrado = ms.ToArray();
                        dataCifradaHex = BitConverter.ToString(bufferCifrado).Replace("-", "");
                    }
                }
                catch (Exception ex)
                {
                    beLog obeLog = new beLog();
                    obeLog.Aplicacion = "Libreria Codigo Usuario - Clase Seguridad";
                    obeLog.MensajeError = ex.Message;
                    obeLog.DetalleError = ex.StackTrace;
                    Log.Grabar(obeLog);
                }
            }
            return dataCifradaHex;
        }

        public static string descifrarAESHex(string mensajeHex)
        {
            byte[] dataCifrada = convertirHexadecimalToBytes(mensajeHex);
            AesCryptoServiceProvider aes = new AesCryptoServiceProvider();
            string strClave = "Curs0P@vn@vMJ".PadRight(32, ' ');
            string strVectorInicio = "Lduenas".PadRight(16, ' ');
            byte[] bufferClave = Encoding.Default.GetBytes(strClave);
            byte[] bufferVectorInicio = Encoding.Default.GetBytes(strVectorInicio);
            string dataDescifrada = "";
            using (MemoryStream ms = new MemoryStream(dataCifrada))
            {
                try
                {
                    using (CryptoStream cs = new CryptoStream(ms, aes.CreateDecryptor(bufferClave, bufferVectorInicio), CryptoStreamMode.Read))
                    {

                        using (StreamReader sr = new StreamReader(cs))
                        {
                            dataDescifrada = sr.ReadToEnd();
                        }
                    }
                }
                catch (Exception ex)
                {
                    beLog obeLog = new beLog();
                    obeLog.Aplicacion = "Libreria Codigo Usuario - Clase Seguridad";
                    obeLog.MensajeError = ex.Message;
                    obeLog.DetalleError = ex.StackTrace;
                    Log.Grabar(obeLog);
                }
            }
            return dataDescifrada;
        }
    }
}
