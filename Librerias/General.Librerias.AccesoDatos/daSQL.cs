using System;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using General.Librerias.CodigoUsuario;

namespace General.Librerias.AccesoDatos
{
    public class daSQL
    {
        public string cadenaConexion;
        public bool esCifrado;
        public daSQL(string NombreConexion, bool cifrado = false)
        {
            esCifrado = cifrado;
            try
            {
                cadenaConexion = ConfigurationManager.ConnectionStrings[NombreConexion].ConnectionString;
                if(cifrado)
                {
                    cadenaConexion = Seguridad.descifrarAESHex(cadenaConexion);
                }
            }
            catch(Exception ex)
            {
                beLog obeLog = new beLog();
                obeLog.MensajeError = ex.Message;
                obeLog.DetalleError = ex.StackTrace;
                Log.Grabar(obeLog, esCifrado);
            }
        }
        public string EjecutarComando(string NombreSP, string ParametroNombre="",string ParametroValor="")
        {
            string rpta = "";
            using (SqlConnection con=new SqlConnection(cadenaConexion))
            {
                try
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(NombreSP, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        if(!string.IsNullOrEmpty(ParametroNombre))
                        {
                            cmd.Parameters.AddWithValue(ParametroNombre, ParametroValor);
                        }
                        object data = cmd.ExecuteScalar();
                        if (data != null) rpta = data.ToString();
                    }
                }
                catch (Exception ex)
                {
                    beLog obeLog = new beLog();
                    obeLog.MensajeError = ex.Message;
                    obeLog.DetalleError = ex.StackTrace;
                    Log.Grabar(obeLog, esCifrado);
                }
            }
            return rpta;
        }

        public byte[] EjecutarComandoBytes(string NombreSP, string ParametroNombre = "", string ParametroValor = "")
        {
            byte[] rpta = null;
            using (SqlConnection con = new SqlConnection(cadenaConexion))
            {
                try
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(NombreSP, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        if (!string.IsNullOrEmpty(ParametroNombre) && !string.IsNullOrEmpty(ParametroValor))
                        {
                            cmd.Parameters.AddWithValue(ParametroNombre, ParametroValor);
                        }
                        object data = cmd.ExecuteScalar();
                        if (data != null) rpta = (byte[])data;
                    }
                }
                catch (Exception ex)
                {
                    beLog obeLog = new beLog();
                    obeLog.MensajeError = ex.Message;
                    obeLog.DetalleError = ex.StackTrace;
                    Log.Grabar(obeLog, esCifrado);
                }
            }
            return rpta;
        }
    }
}
