using System;
using System.IO;
using System.Data;
using System.IO.Compression;
using System.Collections.Generic;
using System.Xml;

namespace General.Librerias.CodigoUsuario
{
    public class ExcelReader
    {
        private static Dictionary<string, string> hojasNombre = new Dictionary<string, string>();
        private static Dictionary<string, string> hojasArchivo = new Dictionary<string, string>();
        private static List<string> listaCadenas = new List<string>();
        private static List<int> indicesFechas = new List<int>();

        private static void iniciarListas()
        {
            hojasNombre = new Dictionary<string, string>();
            hojasArchivo = new Dictionary<string, string>();
            listaCadenas = new List<string>();
            indicesFechas = new List<int>();
        }

        private static string buscarHojaArchivo(string hojaNombre)
        {
            string hojaArchivo = "";
            string hojaId = "";
            foreach (KeyValuePair<string, string> entrada in hojasNombre)
            {
                if (entrada.Value.Equals(hojaNombre))
                {
                    hojaId = entrada.Key;
                    break;
                }
            }
            if (hojaId != "")
            {
                foreach (KeyValuePair<string, string> entrada in hojasArchivo)
                {
                    if (entrada.Key.Equals(hojaId))
                    {
                        hojaArchivo = entrada.Value;
                        break;
                    }
                }
            }
            return hojaArchivo;
        }

        public static List<string> ListarHojas(string archivo)
        {
            List<string> lstHojas = new List<string>();
            XmlDocument doc;
            XmlElement nodoRaizHojas;
            XmlNodeList nodosHojas;
            iniciarListas();
            using (FileStream archivoExcel = new FileStream(archivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var archivoZip = new ZipArchive(archivoExcel, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry archivoXml in archivoZip.Entries)
                {
                    switch (archivoXml.Name)
                    {
                        case "workbook.xml":
                            using (Stream stream = archivoXml.Open())
                            {
                                doc = new XmlDocument();
                                doc.Load(stream);
                                nodoRaizHojas = doc.DocumentElement;
                                nodosHojas = nodoRaizHojas.GetElementsByTagName("sheet");
                                if (nodosHojas != null)
                                {
                                    foreach (XmlNode nodoHoja in nodosHojas)
                                    {
                                        hojasNombre.Add(nodoHoja.Attributes["r:id"].Value, nodoHoja.Attributes["name"].Value);
                                        lstHojas.Add(nodoHoja.Attributes["name"].Value);
                                    }
                                    //lblTotalArchivos.Text = nodosHojas.Count.ToString();
                                }
                            }
                            break;
                        case "workbook.xml.rels":
                            using (Stream stream = archivoXml.Open())
                            {
                                doc = new XmlDocument();
                                doc.Load(stream);
                                nodoRaizHojas = doc.DocumentElement;
                                nodosHojas = nodoRaizHojas.GetElementsByTagName("Relationship");
                                if (nodosHojas != null)
                                {
                                    foreach (XmlNode nodoHoja in nodosHojas)
                                    {
                                        hojasArchivo.Add(nodoHoja.Attributes["Id"].Value, @"xl/" + nodoHoja.Attributes["Target"].Value);
                                    }
                                }
                            }
                            break;
                        case "sharedStrings.xml":
                            using (Stream stream = archivoXml.Open())
                            {
                                using (XmlTextReader xtr = new XmlTextReader(stream))
                                {
                                    while (xtr.Read())
                                    {
                                        if (xtr.Name.Equals("t"))
                                        {
                                            listaCadenas.Add(xtr.ReadString());
                                        }
                                    }
                                }
                            }
                            break;
                        case "styles.xml":
                            string formato;
                            int nFormato;
                            string codigoFormato;
                            using (Stream stream = archivoXml.Open())
                            {
                                using (XmlTextReader xtr = new XmlTextReader(stream))
                                {
                                    while (xtr.Read())
                                    {
                                        if (xtr.Name.Equals("xf"))
                                        {
                                            formato = xtr.GetAttribute("numFmtId");
                                            codigoFormato = xtr.GetAttribute("formatCode");
                                            if (!String.IsNullOrEmpty(formato))
                                            {
                                                nFormato = int.Parse(formato);
                                                if ((nFormato > 13 && nFormato < 23) || (nFormato == 164))
                                                {
                                                    indicesFechas.Add(nFormato);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                    }
                }
            }
            return lstHojas;
        }

        public static DataTable CrearTabla(string nombreArchivo, string nombreHoja)
        {
            DataTable tabla = new DataTable();
            int cCols = 0;
            int cNodos = 0;
            int cFilas = 0;
            string celdaString;
            string celdaEstilo;
            string celdaDimension;
            int posCadena = -1;
            int nFormato = -1;
            bool existeCeldaValor = false;
            bool existeCeldaFormula = false;
            string hojaArchivo = buscarHojaArchivo(nombreHoja);
            string[] campos;
            char sInicio, sFin1;
            char sFin2 = (char)0;
            int n = 0;
            bool esNumero = false;
            bool esCol2 = false;
            DataRow fila = null;
            string valorFecha;
            //DataGridViewTextBoxCell celda = null;
            using (FileStream archivoExcel = new FileStream(nombreArchivo, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var archivoZip = new ZipArchive(archivoExcel, ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry archivoXml in archivoZip.Entries)
                {
                    if (archivoXml.FullName.Equals(hojaArchivo))
                    {
                        using (Stream stream = archivoXml.Open())
                        {
                            using (XmlTextReader xtr = new XmlTextReader(stream))
                            {
                                while (xtr.Read())
                                {
                                    if (xtr.NodeType.Equals(XmlNodeType.Element) && xtr.Name.Equals("dimension"))
                                    {
                                        celdaDimension = xtr.GetAttribute("ref");
                                        if (!String.IsNullOrEmpty(celdaDimension))
                                        {
                                            campos = celdaDimension.Split(':');
                                            esNumero = int.TryParse(campos[0], out n);
                                            sInicio = (campos[0].Substring(0, campos[0].Length - n.ToString().Length))[0];
                                            sFin1 = campos[1][0];
                                            n = (int)(campos[1][1]);
                                            esCol2 = (n >= 65 && n <= 90);
                                            if (esCol2) sFin2 = campos[1][1];
                                            if (!esCol2)
                                            {
                                                for (int j = (int)sInicio; j <= (int)sFin1; j++)
                                                {
                                                    tabla.Columns.Add(((char)j).ToString(), Type.GetType("System.String"));
                                                }
                                            }
                                            else
                                            {
                                                for (int j = (int)sInicio; j <= 90; j++)
                                                {
                                                    tabla.Columns.Add(((char)j).ToString(), Type.GetType("System.String"));
                                                }
                                                for (int j = (int)sInicio; j <= (int)sFin2; j++)
                                                {
                                                    tabla.Columns.Add("A" + ((char)j).ToString(), Type.GetType("System.String"));
                                                }
                                            }
                                        }
                                    }
                                    if (xtr.NodeType.Equals(XmlNodeType.Element) && xtr.Name.Equals("row"))
                                    {
                                        fila = tabla.NewRow();
                                        cCols = 0;
                                    }
                                    if (xtr.NodeType.Equals(XmlNodeType.Element) && xtr.Name.Equals("c"))
                                    {
                                        nFormato = -1;
                                        celdaString = xtr.GetAttribute("t");
                                        celdaEstilo = xtr.GetAttribute("s");
                                        if (!String.IsNullOrEmpty(celdaEstilo))
                                        {
                                            nFormato = int.Parse(celdaEstilo);
                                        }
                                        existeCeldaValor = xtr.ReadToDescendant("v");
                                        if (existeCeldaValor)
                                        {
                                            if (!String.IsNullOrEmpty(celdaString))
                                            {
                                                posCadena = int.Parse(xtr.ReadString());
                                                if (posCadena > -1) fila[cCols] = listaCadenas[posCadena];                                                
                                            }
                                            else
                                            {
                                                if (!String.IsNullOrEmpty(celdaEstilo))
                                                {
                                                    valorFecha = xtr.ReadString();
                                                    fila[cCols] = DateTime.FromOADate(double.Parse(valorFecha)).ToShortDateString();
                                                }
                                                else fila[cCols] = xtr.ReadString();
                                            }
                                        }
                                        existeCeldaFormula = xtr.ReadToDescendant("f");
                                        if (existeCeldaFormula)
                                        {
                                            fila[cCols] = xtr.ReadString();
                                        }
                                        cCols++;
                                    }
                                    if (xtr.NodeType.Equals(XmlNodeType.EndElement) && xtr.Name.Equals("row"))
                                    {
                                        tabla.Rows.Add(fila);
                                        cFilas++;
                                        //if (cFilas == 100000) break;
                                    }
                                    cNodos++;
                                }
                            }
                        }
                    }
                }
            }
            return tabla;
        }
    }
}
