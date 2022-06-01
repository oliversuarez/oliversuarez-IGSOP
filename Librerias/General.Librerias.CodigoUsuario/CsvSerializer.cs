using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Linq;
using System.Reflection;


namespace General.Librerias.CodigoUsuario
{
    public class CsvSerializer
    {
        public static string Serializar<T>(List<T> lista, char separadorCampo, char separadorRegistro, bool incluirCabeceras = true, string archivo = "")
        {
            StringBuilder sb = new StringBuilder();
            PropertyInfo[] propiedades = lista[0].GetType().GetProperties();
            string tipo;
            if (archivo == "")
            {
                if (incluirCabeceras)
                {
                    for (int i = 0; i < propiedades.Length; i++)
                    {
                        sb.Append(propiedades[i].Name);
                        if (i < propiedades.Length - 1) sb.Append(separadorCampo);
                    }
                    sb.Append(separadorRegistro);
                }
                for (int j = 0; j < lista.Count; j++)
                {
                    propiedades = lista[j].GetType().GetProperties();
                    for (int i = 0; i < propiedades.Length; i++)
                    {
                        tipo = propiedades[i].PropertyType.ToString();
                        if (propiedades[i].GetValue(lista[j], null) != null)
                        {
                            if (tipo.Contains("Byte[]"))
                            {
                                byte[] buffer = (byte[])propiedades[i].GetValue(lista[j], null);
                                sb.Append(Convert.ToBase64String(buffer));
                            }
                            else sb.Append(propiedades[i].GetValue(lista[j], null).ToString());
                        }
                        else sb.Append("");
                        if (i < propiedades.Length - 1) sb.Append(separadorCampo);
                    }
                    if (j < lista.Count - 1) sb.Append(separadorRegistro);
                }
            }
            else
            {
                if (File.Exists(archivo))
                {
                    List<string> campos = File.ReadAllLines(archivo).ToList();
                    List<string> props = new List<string>();
                    foreach (PropertyInfo propiedad in propiedades)
                    {
                        props.Add(propiedad.Name);
                    }
                    //Escribir las Cabeceras
                    if (incluirCabeceras)
                    {
                        for (int i = 0; i < campos.Count; i++)
                        {
                            if (props.IndexOf(campos[i]) > -1)
                            {
                                sb.Append(campos[i]);
                                sb.Append(separadorCampo);
                            }
                        }
                        sb = sb.Remove(sb.Length - 1, 1);
                        sb.Append(separadorRegistro);
                    }
                    //Escribir las Filas
                    for (int j = 0; j < lista.Count; j++)
                    {
                        propiedades = lista[j].GetType().GetProperties();
                        for (int i = 0; i < campos.Count; i++)
                        {
                            if (props.IndexOf(campos[i]) > -1)
                            {
                                tipo = lista[j].GetType().GetProperty(campos[i]).PropertyType.ToString();
                                if (lista[j].GetType().GetProperty(campos[i]).GetValue(lista[j], null) != null)
                                {
                                    if (tipo.Contains("Byte[]"))
                                    {
                                        byte[] buffer = (byte[])lista[j].GetType().GetProperty(campos[i]).GetValue(lista[j], null);
                                        sb.Append(Convert.ToBase64String(buffer));
                                    }
                                    else sb.Append(lista[j].GetType().GetProperty(campos[i]).GetValue(lista[j], null).ToString());
                                }
                                else sb.Append("");
                                sb.Append(separadorCampo);
                            }
                        }
                        sb = sb.Remove(sb.Length - 1, 1);
                        sb.Append(separadorRegistro);
                    }
                    sb = sb.Remove(sb.Length - 1, 1);
                }
            }
            return sb.ToString();
        }

        public static void SerializarFast<T>(string archivo, List<T> lista, char separadorCampo, char separadorRegistro, bool incluirCabeceras = true)
        {
            PropertyInfo[] propiedades = lista[0].GetType().GetProperties();
            string tipo;
            using (FileStream fs = new FileStream(archivo, FileMode.Create, FileAccess.Write, FileShare.Write))
            {
                using(StreamWriter sw=new StreamWriter(fs))
                {
                    if (incluirCabeceras)
                    {
                        for (int i = 0; i < propiedades.Length; i++)
                        {
                            sw.Write(propiedades[i].Name);
                            if (i < propiedades.Length - 1) sw.Write(separadorCampo);
                        }
                        sw.Write(separadorRegistro);
                    }
                    for (int j = 0; j < lista.Count; j++)
                    {
                        propiedades = lista[j].GetType().GetProperties();
                        for (int i = 0; i < propiedades.Length; i++)
                        {
                            tipo = propiedades[i].PropertyType.ToString();
                            if (propiedades[i].GetValue(lista[j], null) != null)
                            {
                                if (tipo.Contains("Byte[]"))
                                {
                                    byte[] buffer = (byte[])propiedades[i].GetValue(lista[j], null);
                                    sw.Write(Convert.ToBase64String(buffer));
                                }
                                else sw.Write(propiedades[i].GetValue(lista[j], null).ToString());
                            }
                            else sw.Write("");
                            if (i < propiedades.Length - 1) sw.Write(separadorCampo);
                        }
                        if (j < lista.Count - 1) sw.Write(separadorRegistro);
                    }
                }
            }
        }

        public static List<T> Deserializar<T>(string archivo, char separadorCampo,char separadorRegistro)
        {
            List<T> lista = new List<T>();
            if (File.Exists(archivo))
            {
                string contenido = File.ReadAllText(archivo);
                string[] registros = contenido.Split(separadorRegistro);
                string[] campos;
                string[] cabecera = registros[0].Split(separadorCampo);
                string registro;
                Type tipoObj;
                T obj;
                dynamic valor;
                Type tipoCampo;
                for (int i = 1; i < registros.Length; i++)
                {
                    registro = registros[i];
                    tipoObj = typeof(T);
                    obj = (T)Activator.CreateInstance(tipoObj);
                    campos = registro.Split(separadorCampo);
                    for (int j = 0; j < campos.Length; j++)
                    {
                        tipoCampo = obj.GetType().GetProperty(cabecera[j]).PropertyType;
                        valor = Convert.ChangeType(campos[j], tipoCampo);
                        obj.GetType().GetProperty(cabecera[j]).SetValue(obj, valor);
                    }
                    lista.Add(obj);
                }
            }
            return (lista);
        }

        public static string SerializarObjeto<T>(T obj, char separadorCampo)
        {
            StringBuilder sb = new StringBuilder();
            PropertyInfo[] propiedades = obj.GetType().GetProperties();
            string tipo;
            for (int i = 0; i < propiedades.Length; i++)
            {
                tipo = propiedades[i].PropertyType.ToString();
                if (propiedades[i].GetValue(obj, null) != null)
                {
                    if (tipo.Contains("Byte[]"))
                    {
                        byte[] buffer = (byte[])propiedades[i].GetValue(obj, null);
                        sb.Append(Convert.ToBase64String(buffer));
                    }
                    else sb.Append(propiedades[i].GetValue(obj, null).ToString());
                }
                else sb.Append("");
                if (i < propiedades.Length - 1) sb.Append(separadorCampo);
            }
            return sb.ToString();
        }
    }
}
