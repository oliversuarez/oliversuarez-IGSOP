using System;
using System.Data;

namespace General.Librerias.CodigoUsuario
{
    public class Cadena
    {
        public static DataTable ConvertirTabla(string data, char separadorCampo = '|', char separadorRegistro = '¬')
        {
            DataTable tabla = new DataTable();
            try
            {
                string[] registros = data.Split(separadorRegistro);
                string[] cabeceras = registros[0].Split(separadorCampo);
                string[] anchos = registros[1].Split(separadorCampo);
                string[] tipos = registros[2].Split(separadorCampo);
                int nColumns = cabeceras.Length;
                for (int j = 0; j < nColumns; j++)
                {
                    tabla.Columns.Add(cabeceras[j], Type.GetType("System." + tipos[j]));
                    tabla.Columns[j].Caption = anchos[j];
                }
                int nRegistros = registros.Length;
                string[] campos;
                DataRow fila;
                Type tipo;
                if (nRegistros > 3 && !String.IsNullOrEmpty(registros[3]))
                {
                    for (int i = 3; i < nRegistros; i++)
                    {
                        campos = registros[i].Split(separadorCampo);
                        fila = tabla.NewRow();
                        for (int j = 0; j < nColumns; j++)
                        {
                            tipo = Type.GetType("System." + tipos[j]);
                            fila[j] = Convert.ChangeType(campos[j], tipo);
                        }
                        tabla.Rows.Add(fila);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex.Message);
            }
            return tabla;
        }
    }
}
