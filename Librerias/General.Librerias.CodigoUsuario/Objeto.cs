using System;
using System.IO;
using System.Reflection;

namespace General.Librerias.CodigoUsuario
{
    public class Objeto
    {
        public static void Grabar<T>(T obj,string archivo)
        {
            PropertyInfo[] propiedades = obj.GetType().GetProperties();
            object objValor;
            string strValor;
            using(StreamWriter sw=new StreamWriter(archivo, true))
            {
                foreach (PropertyInfo propiedad in propiedades)
                {
                    strValor = "";
                    objValor = propiedad.GetValue(obj, null);
                    if (objValor != null) strValor = objValor.ToString();
                    sw.WriteLine("{0} = {1}", propiedad.Name, strValor);
                }
                sw.WriteLine(new String('_', 50));
            }
        }
    }
}
