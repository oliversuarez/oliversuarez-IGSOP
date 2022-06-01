﻿using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;

namespace General.Librerias.CodigoUsuario
{
    public class WordMemory
    {
        //private string sArchivoDocx;
        private DataSet dstData;
        private string sTitulo;
        private string sColorFondoCabecera;

        private List<Archivos> LstArchivos = new List<Archivos>();

        public byte[] Exportar(DataSet dst, string titulo = "", string colorFondoCabecera = "BFBFBF")
        {
            //sArchivoDocx = archivoDocx;
            dstData = dst;
            sTitulo = titulo;
            sColorFondoCabecera = colorFondoCabecera;
            return crearDirectoriosArchivos();
        }

        private byte[] crearDirectoriosArchivos()
        {
            //Definir la ruta de los directorios a crear
            //string sDirectorioRaiz = Path.Combine(Path.GetDirectoryName(sArchivoDocx), Path.GetFileNameWithoutExtension(sArchivoDocx));
            //string sDirectorioRels = Path.Combine(sDirectorioRaiz, "_rels");
            //string sDirectorioDocProps = Path.Combine(sDirectorioRaiz, "docProps");
            //string sDirectorioWord = Path.Combine(sDirectorioRaiz, "word");
            //string sDirectorioWordRels = Path.Combine(sDirectorioWord, "_rels");
            //string sDirectorioWordTheme = Path.Combine(sDirectorioWord, "theme");

            //Definir la ruta de los archivos a crear
            Archivos oArchivos = new Archivos();
            //string sArchivoContentTypes = Path.Combine(sDirectorioRaiz, "[Content_Types].xml");
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "[Content_Types]";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoRels = Path.Combine(sDirectorioRels, ".rels");
            oArchivos = new Archivos();
            oArchivos.Extension = "rels";
            oArchivos.Nombre = "_rels/";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoDocApp = Path.Combine(sDirectorioDocProps, "app.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "docProps/app";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoDocCore = Path.Combine(sDirectorioDocProps, "core.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "docProps/core";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordDocument = Path.Combine(sDirectorioWord, "document.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "word/document";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordFontTable = Path.Combine(sDirectorioWord, "fontTable.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "word/fontTable";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordSettings = Path.Combine(sDirectorioWord, "settings.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "word/settings";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordStyles = Path.Combine(sDirectorioWord, "styles.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "word/styles";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordWebSettings = Path.Combine(sDirectorioWord, "webSettings.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "word/webSettings";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordRels = Path.Combine(sDirectorioWordRels, "document.xml.rels");
            oArchivos = new Archivos();
            oArchivos.Extension = "rels";
            oArchivos.Nombre = "word/_rels/document.xml";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //string sArchivoWordTheme = Path.Combine(sDirectorioWordTheme, "theme1.xml");
            oArchivos = new Archivos();
            oArchivos.Extension = "xml";
            oArchivos.Nombre = "word/theme/theme1";
            oArchivos.FileBytes = new byte[] { };
            LstArchivos.Add(oArchivos);

            //Crear los Directorios definidos
            //DirectoryInfo oDirectorioRaiz = Directory.CreateDirectory(sDirectorioRaiz);
            //oDirectorioRaiz.CreateSubdirectory("_rels");
            //oDirectorioRaiz.CreateSubdirectory("docProps");
            //DirectoryInfo oDirectorioWord = oDirectorioRaiz.CreateSubdirectory("word");
            //oDirectorioWord.CreateSubdirectory("_rels");
            //oDirectorioWord.CreateSubdirectory("theme");

            int posArchivo = -1;
            //Crear los Archivos definidos
            //File.WriteAllText(sArchivoContentTypes, getContentTypes());
            posArchivo = buscarArchivo("[Content_Types]");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getContentTypes());
            }

            //File.WriteAllText(sArchivoRels, getRels());
            posArchivo = buscarArchivo("_rels/");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getRels());
            }

            //File.WriteAllText(sArchivoDocApp, getApp());
            posArchivo = buscarArchivo("docProps/app");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getApp());
            }

            //File.WriteAllText(sArchivoDocCore, getCore());
            posArchivo = buscarArchivo("docProps/core");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getCore());
            }

            //File.WriteAllText(sArchivoWordDocument, getWordDocument());
            //getWordDocument(sArchivoWordDocument);
            MemoryStream oMemoryStream;
            posArchivo = buscarArchivo("word/document");
            if (posArchivo > -1)
            {
                oMemoryStream = new MemoryStream();
                getWordDocument(oMemoryStream);
                LstArchivos[posArchivo].FileBytes = oMemoryStream.ToArray();
            }

            //File.WriteAllText(sArchivoWordFontTable, getWordFontTable());
            posArchivo = buscarArchivo("word/fontTable");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getWordFontTable());
            }

            //File.WriteAllText(sArchivoWordSettings, getWordSettings());
            posArchivo = buscarArchivo("word/settings");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getWordSettings());
            }

            //File.WriteAllText(sArchivoWordStyles, getWordStyles());
            posArchivo = buscarArchivo("word/styles");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getWordStyles());
            }


            //File.WriteAllText(sArchivoWordWebSettings, getWordWebSettings());
            posArchivo = buscarArchivo("word/webSettings");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getWordWebSettings());
            }

            //File.WriteAllText(sArchivoWordRels, getWordRels());
            posArchivo = buscarArchivo("word/_rels/document.xml");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getWordRels());
            }

            //File.WriteAllText(sArchivoWordTheme, getWordTheme());
            posArchivo = buscarArchivo("word/theme/theme1");
            if (posArchivo > -1)
            {
                LstArchivos[posArchivo].FileBytes = Encoding.UTF8.GetBytes(getWordTheme());
            }

            //Si el archivo ya existe entonces borrar
            //if (File.Exists(sArchivoDocx)) File.Delete(sArchivoDocx);
            //Comprimir los archivos en un Docx
            //ZipFile.CreateFromDirectory(sDirectorioRaiz, sArchivoDocx);
            //Borrar todo el directorio con los archivos temporales creados
            //Directory.Delete(sDirectorioRaiz, true);

            byte[] fileBytes = null;

            // create a working memory stream
            using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
            {
                // create a zip
                using (System.IO.Compression.ZipArchive zip = new System.IO.Compression.ZipArchive(memoryStream, System.IO.Compression.ZipArchiveMode.Create, true))
                {
                    // interate through the source files
                    foreach (Archivos f in LstArchivos)
                    {
                        // add the item name to the zip
                        System.IO.Compression.ZipArchiveEntry zipItem = zip.CreateEntry(f.Nombre + "." + f.Extension);
                        // add the item bytes to the zip entry by opening the original file and copying the bytes
                        using (System.IO.MemoryStream originalFileMemoryStream = new System.IO.MemoryStream(f.FileBytes))
                        {
                            using (System.IO.Stream entryStream = zipItem.Open())
                            {
                                originalFileMemoryStream.CopyTo(entryStream);
                            }
                        }
                    }
                }
                fileBytes = memoryStream.ToArray();
            }

            return fileBytes;
        }

        private string getContentTypes()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>");
            sb.Append("<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>");
            sb.Append("<Default Extension='xml' ContentType='application/xml'/>");
            sb.Append("<Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>");
            sb.Append("<Override PartName='/word/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'/>");
            sb.Append("<Override PartName='/docProps/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/>");
            sb.Append("<Override PartName='/word/settings.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'/>");
            sb.Append("<Override PartName='/word/theme/theme1.xml' ContentType='application/vnd.openxmlformats-officedocument.theme+xml'/>");
            sb.Append("<Override PartName='/word/fontTable.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml'/>");
            sb.Append("<Override PartName='/word/webSettings.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml'/>");
            sb.Append("<Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml'/>");
            sb.Append("</Types>");
            return sb.ToString();
        }

        private string getRels()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>");
            sb.Append("<Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties' Target='docProps/app.xml'/>");
            sb.Append("<Relationship Id='rId2' Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties' Target='docProps/core.xml'/>");
            sb.Append("<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private string getApp()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'>");
            sb.Append("<Template>Normal.dotm</Template>");
            sb.Append("<TotalTime>5</TotalTime>");
            sb.Append("<Pages>1</Pages>");
            sb.Append("<Words>21</Words>");
            sb.Append("<Characters>116</Characters>");
            sb.Append("<Application>Microsoft Office Word</Application>");
            sb.Append("<DocSecurity>0</DocSecurity>");
            sb.Append("<Lines>1</Lines>");
            sb.Append("<Paragraphs>1</Paragraphs>");
            sb.Append("<ScaleCrop>false</ScaleCrop>");
            sb.Append("<Company>Dsoft</Company>");
            sb.Append("<LinksUpToDate>false</LinksUpToDate>");
            sb.Append("<CharactersWithSpaces>136</CharactersWithSpaces>");
            sb.Append("<SharedDoc>false</SharedDoc>");
            sb.Append("<HyperlinksChanged>false</HyperlinksChanged>");
            sb.Append("<AppVersion>12.0000</AppVersion>");
            sb.Append("</Properties>");
            return sb.ToString();
        }

        private string getCore()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' ");
            sb.Append("xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/' ");
            sb.Append("xmlns:dcmitype='http://purl.org/dc/dcmitype/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>");
            sb.Append("<dc:title></dc:title>");
            sb.Append("<dc:subject></dc:subject>");
            sb.Append("<dc:creator>");
            sb.Append(Environment.UserName);
            sb.Append("</dc:creator>");
            sb.Append("<cp:keywords></cp:keywords>");
            sb.Append("<dc:description></dc:description>");
            sb.Append("<cp:lastModifiedBy>");
            sb.Append(Environment.UserName);
            sb.Append("</cp:lastModifiedBy>");
            sb.Append("<cp:revision>1</cp:revision>");
            sb.Append("<dcterms:created xsi:type='dcterms:W3CDTF'>");
            sb.Append(DateTime.Now.ToString("s"));
            sb.Append("Z</dcterms:created>");
            sb.Append("<dcterms:modified xsi:type='dcterms:W3CDTF'>");
            sb.Append(DateTime.Now.ToString("s"));
            sb.Append("Z</dcterms:modified>");
            sb.Append("</cp:coreProperties>");
            return sb.ToString();
        }

        //private void getWordDocument(string sArchivoWordDocument)
        private void getWordDocument(MemoryStream ms)
        {
            //StringBuilder sb = new StringBuilder();
            //using (StreamWriter sw = new StreamWriter(sArchivoWordDocument))
            using (StreamWriter sw = new StreamWriter(ms))
            {
                sw.Write("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
                sw.Write("<w:document xmlns:ve='http://schemas.openxmlformats.org/markup-compatibility/2006' ");
                sw.Write("xmlns:o='urn:schemas-microsoft-com:office:office' ");
                sw.Write("xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' ");
                sw.Write("xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' ");
                sw.Write("xmlns:v='urn:schemas-microsoft-com:vml' ");
                sw.Write("xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing' ");
                sw.Write("xmlns:w10='urn:schemas-microsoft-com:office:word' ");
                sw.Write("xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ");
                sw.Write("xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'>");

                sw.Write("<w:body>");
                sw.Write("<w:p w:rsidR='0085037B' w:rsidRPr='0045015F' w:rsidRDefault='0045015F' w:rsidP='0045015F'>");
                sw.Write("<w:pPr>");
                sw.Write("<w:jc w:val='center'/>");
                sw.Write("<w:rPr>");
                sw.Write("<w:b/>");
                sw.Write("<w:sz w:val='24'/>");
                sw.Write("</w:rPr>");
                sw.Write("</w:pPr>");
                sw.Write("<w:r w:rsidRPr='0045015F'>");
                sw.Write("<w:rPr>");
                sw.Write("<w:b/>");
                sw.Write("<w:sz w:val='24'/>");
                sw.Write("</w:rPr>");
                sw.Write("<w:t>");
                sw.Write(sTitulo);
                sw.Write("</w:t>");
                sw.Write("</w:r>");
                sw.Write("</w:p>");
                foreach (DataTable tData in dstData.Tables)
                {
                    //Creación de la Tabla
                    sw.Write("<w:tbl>");
                    //Configuración de la Tabla
                    sw.Write("<w:tblPr>");
                    sw.Write("<w:tblStyle w:val='Tablaconcuadrcula'/>");
                    sw.Write("<w:tblW w:w='0' w:type='auto'/>");
                    sw.Write("<w:tblLook w:val='04A0'/>");
                    sw.Write("</w:tblPr>");
                    //Configuración de Columnas
                    sw.Write("<w:tblGrid>");
                    for (int j = 0; j < tData.Columns.Count; j++)
                    {
                        sw.Write("<w:gridCol w:w='2881'/>");
                    }
                    sw.Write("</w:tblGrid>");
                    //Crear la primera fila con las cabeceras
                    sw.Write("<w:tr w:rsidR='0045015F' w:rsidTr='0045015F'>");
                    for (int j = 0; j < tData.Columns.Count; j++)
                    {
                        //Iniciar la creación de la celda
                        sw.Write("<w:tc>");
                        //Configurar la celda
                        sw.Write("<w:tcPr>");
                        sw.Write("<w:tcW w:w='2881' w:type='dxa'/>");
                        sw.Write("<w:shd w:val='clear' w:color='auto' w:fill='");
                        sw.Write(sColorFondoCabecera);
                        sw.Write("'/>");
                        sw.Write("</w:tcPr>");
                        //Crear el contenido de la celda
                        sw.Write("<w:p w:rsidR='0045015F' w:rsidRPr='0045015F' w:rsidRDefault='0045015F' w:rsidP='0045015F'>");
                        sw.Write("<w:r>");
                        //Formatear contenido en Negrita
                        sw.Write("<w:rPr>");
                        sw.Write("<w:b/>");
                        sw.Write("</w:rPr>");
                        sw.Write("<w:t>");
                        sw.Write(tData.Columns[j].ColumnName);
                        sw.Write("</w:t>");
                        sw.Write("</w:r>");
                        sw.Write("</w:p>");
                        //Cerrar la creación de la celda
                        sw.Write("</w:tc>");
                    }
                    sw.Write("</w:tr>");
                    //Crear las filas con los registros de la tabla de datos
                    for (int i = 0; i < tData.Rows.Count; i++)
                    {
                        sw.Write("<w:tr w:rsidR='0045015F' w:rsidTr='0045015F'>");
                        for (int j = 0; j < tData.Columns.Count; j++)
                        {
                            //Iniciar la creación de la celda
                            sw.Write("<w:tc>");
                            //Configurar la celda
                            sw.Write("<w:tcPr>");
                            sw.Write("<w:tcW w:w='2881' w:type='dxa'/>");
                            sw.Write("</w:tcPr>");
                            //Crear el contenido de la celda
                            sw.Write("<w:p w:rsidR='0045015F' w:rsidRDefault='0045015F'>");
                            sw.Write("<w:r>");
                            sw.Write("<w:t>");
                            if (tData.Columns[j].DataType.ToString().Contains("Date")) sw.Write(String.Format("{0:dd/MM/yyyy}", tData.Rows[i][j]));
                            else sw.Write(tData.Rows[i][j].ToString());
                            sw.Write("</w:t>");
                            sw.Write("</w:r>");
                            sw.Write("</w:p>");
                            //Cerrar la creación de la celda
                            sw.Write("</w:tc>");
                        }
                        sw.Write("</w:tr>");
                    }
                    sw.Write("</w:tbl>");
                }
                sw.Write("<w:p w:rsidR='0045015F' w:rsidRDefault='0045015F'/>");
                sw.Write("<w:sectPr w:rsidR='0045015F' w:rsidSect='003D3FA0'>");
                sw.Write("<w:pgSz w:w='11906' w:h='16838'/>");
                sw.Write("<w:pgMar w:top='1417' w:right='1701' w:bottom='1417' w:left='1701' w:header='708' w:footer='708' w:gutter='0'/>");
                sw.Write("<w:cols w:space='708'/>");
                sw.Write("<w:docGrid w:linePitch='360'/>");
                sw.Write("</w:sectPr>");
                sw.Write("</w:body>");
                sw.Write("</w:document>");
            }
            //return sb.ToString();
        }

        private string getWordFontTable()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<w:fonts xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' ");
            sb.Append("xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>");
            sb.Append("<w:font w:name='Calibri'>");
            sb.Append("<w:panose1 w:val='020F0502020204030204'/>");
            sb.Append("<w:charset w:val='00'/>");
            sb.Append("<w:family w:val='swiss'/>");
            sb.Append("<w:pitch w:val='variable'/>");
            sb.Append("<w:sig w:usb0='E00002FF' w:usb1='4000ACFF' ");
            sb.Append("w:usb2='00000001' w:usb3='00000000' ");
            sb.Append("w:csb0='0000019F' w:csb1='00000000'/>");
            sb.Append("</w:font>");
            sb.Append("<w:font w:name='Times New Roman'>");
            sb.Append("<w:panose1 w:val='02020603050405020304'/>");
            sb.Append("<w:charset w:val='00'/>");
            sb.Append("<w:family w:val='roman'/>");
            sb.Append("<w:pitch w:val='variable'/>");
            sb.Append("<w:sig w:usb0='E0002EFF' w:usb1='C0007843' ");
            sb.Append("w:usb2='00000009' w:usb3='00000000' ");
            sb.Append("w:csb0='000001FF' w:csb1='00000000'/>");
            sb.Append("</w:font>");
            sb.Append("<w:font w:name='Calibri Light'>");
            sb.Append("<w:panose1 w:val='020F0302020204030204'/>");
            sb.Append("<w:charset w:val='00'/>");
            sb.Append("<w:family w:val='swiss'/>");
            sb.Append("<w:pitch w:val='variable'/>");
            sb.Append("<w:sig w:usb0='A00002EF' w:usb1='4000207B' ");
            sb.Append("w:usb2='00000000' w:usb3='00000000' ");
            sb.Append("w:csb0='0000019F' w:csb1='00000000'/>");
            sb.Append("</w:font>");
            sb.Append("</w:fonts>");
            return sb.ToString();
        }

        private string getWordSettings()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<w:settings xmlns:o='urn:schemas-microsoft-com:office:office' ");
            sb.Append("xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' ");
            sb.Append("xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math' ");
            sb.Append("xmlns:v='urn:schemas-microsoft-com:vml' ");
            sb.Append("xmlns:w10='urn:schemas-microsoft-com:office:word' ");
            sb.Append("xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' ");
            sb.Append("xmlns:sl='http://schemas.openxmlformats.org/schemaLibrary/2006/main'>");
            sb.Append("<w:zoom w:percent='100'/>");
            sb.Append("<w:proofState w:spelling='clean' w:grammar='clean'/>");
            sb.Append("<w:defaultTabStop w:val='708'/>");
            sb.Append("<w:hyphenationZone w:val='425'/>");
            sb.Append("<w:characterSpacingControl w:val='doNotCompress'/>");
            sb.Append("<w:compat/>");
            sb.Append("<w:rsids>");
            sb.Append("<w:rsidRoot w:val='0045015F'/>");
            sb.Append("<w:rsid w:val='00355FE2'/>");
            sb.Append("<w:rsid w:val='003D3FA0'/>");
            sb.Append("<w:rsid w:val='00435B6B'/>");
            sb.Append("<w:rsid w:val='0045015F'/>");
            sb.Append("<w:rsid w:val='004A57D2'/>");
            sb.Append("<w:rsid w:val='0085037B'/>");
            sb.Append("<w:rsid w:val='00854D17'/>");
            sb.Append("<w:rsid w:val='00AC0ECB'/>");
            sb.Append("<w:rsid w:val='00E962BE'/>");
            sb.Append("<w:rsid w:val='00EC1EF3'/>");
            sb.Append("</w:rsids>");
            sb.Append("<m:mathPr>");
            sb.Append("<m:mathFont m:val='Cambria Math'/>");
            sb.Append("<m:brkBin m:val='before'/>");
            sb.Append("<m:brkBinSub m:val='--'/>");
            sb.Append("<m:smallFrac/>");
            sb.Append("<m:dispDef/>");
            sb.Append("<m:lMargin m:val='0'/>");
            sb.Append("<m:rMargin m:val='0'/>");
            sb.Append("<m:defJc m:val='centerGroup'/>");
            sb.Append("<m:wrapIndent m:val='1440'/>");
            sb.Append("<m:intLim m:val='subSup'/>");
            sb.Append("<m:naryLim m:val='undOvr'/>");
            sb.Append("</m:mathPr>");
            sb.Append("<w:themeFontLang w:val='es-PE'/>");
            sb.Append("<w:clrSchemeMapping w:bg1='light1' w:t1='dark1' ");
            sb.Append("w:bg2='light2' w:t2='dark2' w:accent1='accent1' ");
            sb.Append("w:accent2='accent2' w:accent3='accent3' ");
            sb.Append("w:accent4='accent4' w:accent5='accent5' ");
            sb.Append("w:accent6='accent6' w:hyperlink='hyperlink' ");
            sb.Append("w:followedHyperlink='followedHyperlink'/>");
            sb.Append("<w:shapeDefaults>");
            sb.Append("<o:shapedefaults v:ext='edit' spidmax='1026'/>");
            sb.Append("<o:shapelayout v:ext='edit'>");
            sb.Append("<o:idmap v:ext='edit' data='1'/>");
            sb.Append("</o:shapelayout>");
            sb.Append("</w:shapeDefaults>");
            sb.Append("<w:decimalSymbol w:val='.'/>");
            sb.Append("<w:listSeparator w:val=','/>");
            sb.Append("</w:settings>");
            return sb.ToString();
        }

        private string getWordStyles()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<w:styles xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' ");
            sb.Append("xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>");
            sb.Append("<w:docDefaults>");
            sb.Append("<w:rPrDefault>");
            sb.Append("<w:rPr>");
            sb.Append("<w:rFonts w:asciiTheme='minorHAnsi' ");
            sb.Append("w:eastAsiaTheme='minorHAnsi' ");
            sb.Append("w:hAnsiTheme='minorHAnsi' ");
            sb.Append("w:cstheme='minorBidi'/>");
            sb.Append("<w:sz w:val='22'/>");
            sb.Append("<w:szCs w:val='22'/>");
            sb.Append("<w:lang w:val='es-PE' w:eastAsia='en-US' w:bidi='ar-SA'/>");
            sb.Append("</w:rPr>");
            sb.Append("</w:rPrDefault>");
            sb.Append("<w:pPrDefault>");
            sb.Append("<w:pPr>");
            sb.Append("<w:spacing w:after='160' w:line='259' w:lineRule='auto'/>");
            sb.Append("</w:pPr>");
            sb.Append("</w:pPrDefault>");
            sb.Append("</w:docDefaults>");
            sb.Append("<w:latentStyles w:defLockedState='0' w:defUIPriority='99' ");
            sb.Append("w:defSemiHidden='1' w:defUnhideWhenUsed='1' ");
            sb.Append("w:defQFormat='0' w:count='267'>");
            sb.Append("<w:lsdException w:name='Normal' w:semiHidden='0' ");
            sb.Append("w:uiPriority='0' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 1' ");
            sb.Append("w:semiHidden='0' w:uiPriority='9' ");
            sb.Append("w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 2' ");
            sb.Append("w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 3' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 4' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 5' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 6' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 7' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 8' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='heading 9' w:uiPriority='9' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='toc 1' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 2' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 3' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 4' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 5' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 6' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 7' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 8' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='toc 9' w:uiPriority='39'/>");
            sb.Append("<w:lsdException w:name='caption' w:uiPriority='35' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Title' w:semiHidden='0' ");
            sb.Append("w:uiPriority='10' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Default Paragraph Font' w:uiPriority='1'/>");
            sb.Append("<w:lsdException w:name='Subtitle' w:semiHidden='0' w:uiPriority='11' ");
            sb.Append("w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Strong' w:semiHidden='0' ");
            sb.Append("w:uiPriority='22' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Emphasis' w:semiHidden='0' ");
            sb.Append("w:uiPriority='20' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Table Grid' w:semiHidden='0' ");
            sb.Append("w:uiPriority='39' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Placeholder Text' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='No Spacing' w:semiHidden='0' ");
            sb.Append("w:uiPriority='1' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Light Shading' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Shading Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Revision' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='List Paragraph' w:semiHidden='0' ");
            sb.Append("w:uiPriority='34' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Quote' w:semiHidden='0' w:uiPriority='29' ");
            sb.Append("w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Intense Quote' w:semiHidden='0' ");
            sb.Append("w:uiPriority='30' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Medium List 2 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3 Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid Accent 1' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Shading Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 2 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3 Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid Accent 2' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Shading Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 2 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3 Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid Accent 3' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Shading Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 2 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3 Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid Accent 4' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Shading Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 2 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3 Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid Accent 5' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Shading Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='60' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light List Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='61' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Light Grid Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='62' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 1 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='63' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Shading 2 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='64' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 1 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='65' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium List 2 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='66' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 1 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='67' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 2 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='68' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Medium Grid 3 Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='69' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Dark List Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='70' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Shading Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='71' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful List Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='72' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Colorful Grid Accent 6' w:semiHidden='0' ");
            sb.Append("w:uiPriority='73' w:unhideWhenUsed='0'/>");
            sb.Append("<w:lsdException w:name='Subtle Emphasis' w:semiHidden='0' ");
            sb.Append("w:uiPriority='19' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Intense Emphasis' w:semiHidden='0' ");
            sb.Append("w:uiPriority='21' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Subtle Reference' w:semiHidden='0' ");
            sb.Append("w:uiPriority='31' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Intense Reference' w:semiHidden='0' ");
            sb.Append("w:uiPriority='32' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Book Title' w:semiHidden='0' ");
            sb.Append("w:uiPriority='33' w:unhideWhenUsed='0' w:qFormat='1'/>");
            sb.Append("<w:lsdException w:name='Bibliography' w:uiPriority='37'/>");
            sb.Append("<w:lsdException w:name='TOC Heading' w:uiPriority='39' w:qFormat='1'/>");
            sb.Append("</w:latentStyles>");
            sb.Append("<w:style w:type='paragraph' w:default='1' w:styleId='Normal'>");
            sb.Append("<w:name w:val='Normal'/>");
            sb.Append("<w:qFormat/>");
            sb.Append("<w:rsid w:val='003D3FA0'/>");
            sb.Append("</w:style>");
            sb.Append("<w:style w:type='character' w:default='1' w:styleId='Fuentedeprrafopredeter'>");
            sb.Append("<w:name w:val='Default Paragraph Font'/>");
            sb.Append("<w:uiPriority w:val='1'/>");
            sb.Append("<w:semiHidden/>");
            sb.Append("<w:unhideWhenUsed/>");
            sb.Append("</w:style>");
            sb.Append("<w:style w:type='table' w:default='1' w:styleId='Tablanormal'>");
            sb.Append("<w:name w:val='Normal Table'/>");
            sb.Append("<w:uiPriority w:val='99'/>");
            sb.Append("<w:semiHidden/>");
            sb.Append("<w:unhideWhenUsed/>");
            sb.Append("<w:qFormat/>");
            sb.Append("<w:tblPr>");
            sb.Append("<w:tblInd w:w='0' w:type='dxa'/>");
            sb.Append("<w:tblCellMar>");
            sb.Append("<w:top w:w='0' w:type='dxa'/>");
            sb.Append("<w:left w:w='108' w:type='dxa'/>");
            sb.Append("<w:bottom w:w='0' w:type='dxa'/>");
            sb.Append("<w:right w:w='108' w:type='dxa'/>");
            sb.Append("</w:tblCellMar>");
            sb.Append("</w:tblPr>");
            sb.Append("</w:style>");
            sb.Append("<w:style w:type='numbering' w:default='1' w:styleId='Sinlista'>");
            sb.Append("<w:name w:val='No List'/>");
            sb.Append("<w:uiPriority w:val='99'/>");
            sb.Append("<w:semiHidden/>");
            sb.Append("<w:unhideWhenUsed/>");
            sb.Append("</w:style>");
            sb.Append("<w:style w:type='table' w:styleId='Tablaconcuadrcula'>");
            sb.Append("<w:name w:val='Table Grid'/>");
            sb.Append("<w:basedOn w:val='Tablanormal'/>");
            sb.Append("<w:uiPriority w:val='39'/>");
            sb.Append("<w:rsid w:val='0045015F'/>");
            sb.Append("<w:pPr>");
            sb.Append("<w:spacing w:after='0' w:line='240' w:lineRule='auto'/>");
            sb.Append("</w:pPr>");
            sb.Append("<w:tblPr>");
            sb.Append("<w:tblInd w:w='0' w:type='dxa'/>");
            sb.Append("<w:tblBorders>");
            sb.Append("<w:top w:val='single' w:sz='4' w:space='0' w:color='auto'/>");
            sb.Append("<w:left w:val='single' w:sz='4' w:space='0' w:color='auto'/>");
            sb.Append("<w:bottom w:val='single' w:sz='4' w:space='0' w:color='auto'/>");
            sb.Append("<w:right w:val='single' w:sz='4' w:space='0' w:color='auto'/>");
            sb.Append("<w:insideH w:val='single' w:sz='4' w:space='0' w:color='auto'/>");
            sb.Append("<w:insideV w:val='single' w:sz='4' w:space='0' w:color='auto'/>");
            sb.Append("</w:tblBorders>");
            sb.Append("<w:tblCellMar>");
            sb.Append("<w:top w:w='0' w:type='dxa'/>");
            sb.Append("<w:left w:w='108' w:type='dxa'/>");
            sb.Append("<w:bottom w:w='0' w:type='dxa'/>");
            sb.Append("<w:right w:w='108' w:type='dxa'/>");
            sb.Append("</w:tblCellMar>");
            sb.Append("</w:tblPr>");
            sb.Append("</w:style>");
            sb.Append("</w:styles>");
            return sb.ToString();
        }

        private string getWordWebSettings()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<w:webSettings xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' ");
            sb.Append("xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>");
            sb.Append("<w:optimizeForBrowser/>");
            sb.Append("</w:webSettings>");
            return sb.ToString();
        }

        private string getWordRels()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>");
            sb.Append("<Relationship Id='rId3' ");
            sb.Append("Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings' ");
            sb.Append("Target='webSettings.xml'/>");
            sb.Append("<Relationship Id='rId2' ");
            sb.Append("Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings' ");
            sb.Append("Target='settings.xml'/>");
            sb.Append("<Relationship Id='rId1' ");
            sb.Append("Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' ");
            sb.Append("Target='styles.xml'/>");
            sb.Append("<Relationship Id='rId5' ");
            sb.Append("Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' ");
            sb.Append("Target='theme/theme1.xml'/>");
            sb.Append("<Relationship Id='rId4' ");
            sb.Append("Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable' ");
            sb.Append("Target='fontTable.xml'/>");
            sb.Append("</Relationships>");
            return sb.ToString();
        }

        private string getWordTheme()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<?xml version='1.0' encoding='UTF-8' standalone='yes'?>");
            sb.Append("<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' ");
            sb.Append("name='Tema de Office'>");
            sb.Append("<a:themeElements>");
            sb.Append("<a:clrScheme name='Office'>");
            sb.Append("<a:dk1>");
            sb.Append("<a:sysClr val='windowText' lastClr='000000'/>");
            sb.Append("</a:dk1>");
            sb.Append("<a:lt1>");
            sb.Append("<a:sysClr val='window' lastClr='FFFFFF'/>");
            sb.Append("</a:lt1>");
            sb.Append("<a:dk2>");
            sb.Append("<a:srgbClr val='44546A'/>");
            sb.Append("</a:dk2>");
            sb.Append("<a:lt2>");
            sb.Append("<a:srgbClr val='E7E6E6'/>");
            sb.Append("</a:lt2>");
            sb.Append("<a:accent1>");
            sb.Append("<a:srgbClr val='5B9BD5'/>");
            sb.Append("</a:accent1>");
            sb.Append("<a:accent2>");
            sb.Append("<a:srgbClr val='ED7D31'/>");
            sb.Append("</a:accent2>");
            sb.Append("<a:accent3>");
            sb.Append("<a:srgbClr val='A5A5A5'/>");
            sb.Append("</a:accent3>");
            sb.Append("<a:accent4>");
            sb.Append("<a:srgbClr val='FFC000'/>");
            sb.Append("</a:accent4>");
            sb.Append("<a:accent5>");
            sb.Append("<a:srgbClr val='4472C4'/>");
            sb.Append("</a:accent5>");
            sb.Append("<a:accent6>");
            sb.Append("<a:srgbClr val='70AD47'/>");
            sb.Append("</a:accent6>");
            sb.Append("<a:hlink>");
            sb.Append("<a:srgbClr val='0563C1'/>");
            sb.Append("</a:hlink>");
            sb.Append("<a:folHlink>");
            sb.Append("<a:srgbClr val='954F72'/>");
            sb.Append("</a:folHlink>");
            sb.Append("</a:clrScheme>");
            sb.Append("<a:fontScheme name='Office'>");
            sb.Append("<a:majorFont>");
            sb.Append("<a:latin typeface='Calibri Light'/>");
            sb.Append("<a:ea typeface=''/>");
            sb.Append("<a:cs typeface=''/>");
            sb.Append("<a:font script='Jpan' typeface='ＭＳ ゴシック'/>");
            sb.Append("<a:font script='Hang' typeface='맑은 고딕'/>");
            sb.Append("<a:font script='Hans' typeface='宋体'/>");
            sb.Append("<a:font script='Hant' typeface='新細明體'/>");
            sb.Append("<a:font script='Arab' typeface='Times New Roman'/>");
            sb.Append("<a:font script='Hebr' typeface='Times New Roman'/>");
            sb.Append("<a:font script='Thai' typeface='Angsana New'/>");
            sb.Append("<a:font script='Ethi' typeface='Nyala'/>");
            sb.Append("<a:font script='Beng' typeface='Vrinda'/>");
            sb.Append("<a:font script='Gujr' typeface='Shruti'/>");
            sb.Append("<a:font script='Khmr' typeface='MoolBoran'/>");
            sb.Append("<a:font script='Knda' typeface='Tunga'/>");
            sb.Append("<a:font script='Guru' typeface='Raavi'/>");
            sb.Append("<a:font script='Cans' typeface='Euphemia'/>");
            sb.Append("<a:font script='Cher' typeface='Plantagenet Cherokee'/>");
            sb.Append("<a:font script='Yiii' typeface='Microsoft Yi Baiti'/>");
            sb.Append("<a:font script='Tibt' typeface='Microsoft Himalaya'/>");
            sb.Append("<a:font script='Thaa' typeface='MV Boli'/>");
            sb.Append("<a:font script='Deva' typeface='Mangal'/>");
            sb.Append("<a:font script='Telu' typeface='Gautami'/>");
            sb.Append("<a:font script='Taml' typeface='Latha'/>");
            sb.Append("<a:font script='Syrc' typeface='Estrangelo Edessa'/>");
            sb.Append("<a:font script='Orya' typeface='Kalinga'/>");
            sb.Append("<a:font script='Mlym' typeface='Kartika'/>");
            sb.Append("<a:font script='Laoo' typeface='DokChampa'/>");
            sb.Append("<a:font script='Sinh' typeface='Iskoola Pota'/>");
            sb.Append("<a:font script='Mong' typeface='Mongolian Baiti'/>");
            sb.Append("<a:font script='Viet' typeface='Times New Roman'/>");
            sb.Append("<a:font script='Uigh' typeface='Microsoft Uighur'/>");
            sb.Append("<a:font script='Geor' typeface='Sylfaen'/>");
            sb.Append("</a:majorFont>");
            sb.Append("<a:minorFont>");
            sb.Append("<a:latin typeface='Calibri'/>");
            sb.Append("<a:ea typeface=''/>");
            sb.Append("<a:cs typeface=''/>");
            sb.Append("<a:font script='Jpan' typeface='ＭＳ 明朝'/>");
            sb.Append("<a:font script='Hang' typeface='맑은 고딕'/>");
            sb.Append("<a:font script='Hans' typeface='宋体'/>");
            sb.Append("<a:font script='Hant' typeface='新細明體'/>");
            sb.Append("<a:font script='Arab' typeface='Arial'/>");
            sb.Append("<a:font script='Hebr' typeface='Arial'/>");
            sb.Append("<a:font script='Thai' typeface='Cordia New'/>");
            sb.Append("<a:font script='Ethi' typeface='Nyala'/>");
            sb.Append("<a:font script='Beng' typeface='Vrinda'/>");
            sb.Append("<a:font script='Gujr' typeface='Shruti'/>");
            sb.Append("<a:font script='Khmr' typeface='DaunPenh'/>");
            sb.Append("<a:font script='Knda' typeface='Tunga'/>");
            sb.Append("<a:font script='Guru' typeface='Raavi'/>");
            sb.Append("<a:font script='Cans' typeface='Euphemia'/>");
            sb.Append("<a:font script='Cher' typeface='Plantagenet Cherokee'/>");
            sb.Append("<a:font script='Yiii' typeface='Microsoft Yi Baiti'/>");
            sb.Append("<a:font script='Tibt' typeface='Microsoft Himalaya'/>");
            sb.Append("<a:font script='Thaa' typeface='MV Boli'/>");
            sb.Append("<a:font script='Deva' typeface='Mangal'/>");
            sb.Append("<a:font script='Telu' typeface='Gautami'/>");
            sb.Append("<a:font script='Taml' typeface='Latha'/>");
            sb.Append("<a:font script='Syrc' typeface='Estrangelo Edessa'/>");
            sb.Append("<a:font script='Orya' typeface='Kalinga'/>");
            sb.Append("<a:font script='Mlym' typeface='Kartika'/>");
            sb.Append("<a:font script='Laoo' typeface='DokChampa'/>");
            sb.Append("<a:font script='Sinh' typeface='Iskoola Pota'/>");
            sb.Append("<a:font script='Mong' typeface='Mongolian Baiti'/>");
            sb.Append("<a:font script='Viet' typeface='Arial'/>");
            sb.Append("<a:font script='Uigh' typeface='Microsoft Uighur'/>");
            sb.Append("<a:font script='Geor' typeface='Sylfaen'/>");
            sb.Append("</a:minorFont>");
            sb.Append("</a:fontScheme>");
            sb.Append("<a:fmtScheme name='Office'>");
            sb.Append("<a:fillStyleLst>");
            sb.Append("<a:solidFill>");
            sb.Append("<a:schemeClr val='phClr'/>");
            sb.Append("</a:solidFill>");
            sb.Append("<a:gradFill rotWithShape='1'>");
            sb.Append("<a:gsLst>");
            sb.Append("<a:gs pos='0'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:lumMod val='110000'/>");
            sb.Append("<a:satMod val='105000'/>");
            sb.Append("<a:tint val='67000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("<a:gs pos='50000'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:lumMod val='105000'/>");
            sb.Append("<a:satMod val='103000'/>");
            sb.Append("<a:tint val='73000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("<a:gs pos='100000'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:lumMod val='105000'/>");
            sb.Append("<a:satMod val='109000'/>");
            sb.Append("<a:tint val='81000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("</a:gsLst>");
            sb.Append("<a:lin ang='5400000' scaled='0'/>");
            sb.Append("</a:gradFill>");
            sb.Append("<a:gradFill rotWithShape='1'>");
            sb.Append("<a:gsLst>");
            sb.Append("<a:gs pos='0'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:satMod val='103000'/>");
            sb.Append("<a:lumMod val='102000'/>");
            sb.Append("<a:tint val='94000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("<a:gs pos='50000'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:satMod val='110000'/>");
            sb.Append("<a:lumMod val='100000'/>");
            sb.Append("<a:shade val='100000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("<a:gs pos='100000'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:lumMod val='99000'/>");
            sb.Append("<a:satMod val='120000'/>");
            sb.Append("<a:shade val='78000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("</a:gsLst>");
            sb.Append("<a:lin ang='5400000' scaled='0'/>");
            sb.Append("</a:gradFill>");
            sb.Append("</a:fillStyleLst>");
            sb.Append("<a:lnStyleLst>");
            sb.Append("<a:ln w='6350' cap='flat' cmpd='sng' algn='ctr'>");
            sb.Append("<a:solidFill>");
            sb.Append("<a:schemeClr val='phClr'/>");
            sb.Append("</a:solidFill>");
            sb.Append("<a:prstDash val='solid'/>");
            sb.Append("<a:miter lim='800000'/>");
            sb.Append("</a:ln>");
            sb.Append("<a:ln w='12700' cap='flat' cmpd='sng' algn='ctr'>");
            sb.Append("<a:solidFill>");
            sb.Append("<a:schemeClr val='phClr'/>");
            sb.Append("</a:solidFill>");
            sb.Append("<a:prstDash val='solid'/>");
            sb.Append("<a:miter lim='800000'/>");
            sb.Append("</a:ln>");
            sb.Append("<a:ln w='19050' cap='flat' cmpd='sng' algn='ctr'>");
            sb.Append("<a:solidFill>");
            sb.Append("<a:schemeClr val='phClr'/>");
            sb.Append("</a:solidFill>");
            sb.Append("<a:prstDash val='solid'/>");
            sb.Append("<a:miter lim='800000'/>");
            sb.Append("</a:ln>");
            sb.Append("</a:lnStyleLst>");
            sb.Append("<a:effectStyleLst>");
            sb.Append("<a:effectStyle>");
            sb.Append("<a:effectLst/>");
            sb.Append("</a:effectStyle>");
            sb.Append("<a:effectStyle>");
            sb.Append("<a:effectLst/>");
            sb.Append("</a:effectStyle>");
            sb.Append("<a:effectStyle>");
            sb.Append("<a:effectLst>");
            sb.Append("<a:outerShdw blurRad='57150' dist='19050' dir='5400000' algn='ctr' rotWithShape='0'>");
            sb.Append("<a:srgbClr val='000000'>");
            sb.Append("<a:alpha val='63000'/>");
            sb.Append("</a:srgbClr>");
            sb.Append("</a:outerShdw>");
            sb.Append("</a:effectLst>");
            sb.Append("</a:effectStyle>");
            sb.Append("</a:effectStyleLst>");
            sb.Append("<a:bgFillStyleLst>");
            sb.Append("<a:solidFill>");
            sb.Append("<a:schemeClr val='phClr'/>");
            sb.Append("</a:solidFill>");
            sb.Append("<a:solidFill>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:tint val='95000'/>");
            sb.Append("<a:satMod val='170000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:solidFill>");
            sb.Append("<a:gradFill rotWithShape='1'>");
            sb.Append("<a:gsLst>");
            sb.Append("<a:gs pos='0'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:tint val='93000'/>");
            sb.Append("<a:satMod val='150000'/>");
            sb.Append("<a:shade val='98000'/>");
            sb.Append("<a:lumMod val='102000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("<a:gs pos='50000'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:tint val='98000'/>");
            sb.Append("<a:satMod val='130000'/>");
            sb.Append("<a:shade val='90000'/>");
            sb.Append("<a:lumMod val='103000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("<a:gs pos='100000'>");
            sb.Append("<a:schemeClr val='phClr'>");
            sb.Append("<a:shade val='63000'/>");
            sb.Append("<a:satMod val='120000'/>");
            sb.Append("</a:schemeClr>");
            sb.Append("</a:gs>");
            sb.Append("</a:gsLst>");
            sb.Append("<a:lin ang='5400000' scaled='0'/>");
            sb.Append("</a:gradFill>");
            sb.Append("</a:bgFillStyleLst>");
            sb.Append("</a:fmtScheme>");
            sb.Append("</a:themeElements>");
            sb.Append("<a:objectDefaults/>");
            sb.Append("<a:extraClrSchemeLst/>");
            sb.Append("<a:extLst>");
            sb.Append("<a:ext uri='{05A4C25C-085E-4340-85A3-A5531E510DB2}'>");
            sb.Append("<thm15:themeFamily xmlns:thm15='http://schemas.microsoft.com/office/thememl/2012/main' ");
            sb.Append("xmlns='' name='Office Theme' id='{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}' ");
            sb.Append("vid='{4A3C46E8-61CC-4603-A589-7422A47A8E4A}'/>");
            sb.Append("</a:ext>");
            sb.Append("</a:extLst>");
            sb.Append("</a:theme>");
            return sb.ToString();
        }

        private int buscarArchivo(string nombre)
        {
            int narchivos = LstArchivos.Count;
            Archivos oArchivo;
            int pos = -1;
            for (int i = 0; i < narchivos; i++)
            {
                oArchivo = LstArchivos[i];
                if (oArchivo.Nombre.Equals(nombre))
                {
                    pos = i;
                    break;
                }
            }
            return pos;
        }
    }

}
