using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using iText.IO.Font.Constants;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Properties;

namespace General.Librerias.CodigoUsuario
{
    public class PDF
    {
        private static Table serializarTablaPDF(Document oDocumento, DataTable tabla, string titulo, bool flagUltimaCol = false)
        {
            int nCols = tabla.Columns.Count;
            int nFilas = tabla.Rows.Count;
            string campo;
            string ancho;
            List<int> anchos = new List<int>();
            Table tablaPdf = new Table(nCols).UseAllAvailableWidth();

            string Titulo = titulo;
            PdfFont fuente = PdfFontFactory.CreateFont(StandardFonts.TIMES_ROMAN);

            Style styleTitulo = new Style()
                .SetTextAlignment(TextAlignment.CENTER)
                .SetBorder(Border.NO_BORDER)
                .SetBold()
                .SetFont(fuente);

            Style styleCellCabecera = new Style()
                .SetBackgroundColor(ColorConstants.LIGHT_GRAY)
                .SetTextAlignment(TextAlignment.CENTER)
                .SetFontSize(9)
                .SetFont(fuente)
                .SetBold();
            Style styleCellDetalle = new Style()
                .SetBackgroundColor(ColorConstants.WHITE)
                .SetFontSize(9)
                .SetFont(fuente);

            Cell _cell = new Cell(1, nCols).Add(new Paragraph(Titulo));
            tablaPdf.AddHeaderCell(_cell.AddStyle(styleTitulo));

            if (nCols > 0 && nFilas > 0)
            {
                for (int j = 0; j < nCols; j++)
                {
                    campo = tabla.Columns[j].ColumnName;
                    ancho = tabla.Columns[j].Caption;
                    anchos.Add(int.Parse(ancho));
                    _cell = new Cell(j, 1).Add(new Paragraph(campo.ToUpper()));
                    tablaPdf.AddHeaderCell(_cell.AddStyle(styleCellCabecera));
                }
                for (int i = 0; i < nFilas; i++)
                {
                    for (int j = 0; j < nCols; j++)
                    {
                        if (tabla.Rows[i][j].ToString() == "0" || tabla.Rows[i][j].ToString() == "01/01/1900")
                        {
                            _cell = new Cell().Add(new Paragraph(""));
                        }
                        else
                        {
                            _cell = new Cell().Add(new Paragraph(tabla.Rows[i][j].ToString()));
                        }
                        if (flagUltimaCol && tabla.Rows[i][nCols].Equals("1"))
                        {
                            _cell = new Cell().Add(new Paragraph(tabla.Rows[i][j].ToString()));
                            _cell.SetBackgroundColor(ColorConstants.LIGHT_GRAY);
                        }
                        _cell.SetWidth(anchos[j]);
                        tablaPdf.AddCell(_cell.AddStyle(styleCellDetalle));
                    }
                }
            }
            return tablaPdf;
        }

        public static byte[] Crear(DataSet dst, string Titulo = "",string usuario = "",string logo="", string orienta = "", string imagenBase64 = "", int imagenAncho = 0, int imagenAlto = 0)
        {
            byte[] buffer = null;
            using (MemoryStream ms = new MemoryStream())
            {
                PdfWriter pw = new PdfWriter(ms);
                PdfDocument pdfDocument = new PdfDocument(pw);
                Document doc = new Document(pdfDocument);

                if (!String.IsNullOrEmpty(orienta) && orienta == "V")
                {
                    doc = new Document(pdfDocument, PageSize.A4);
                    doc.SetMargins(75, 20, 5, 20);
                }
                else
                {
                    doc = new Document(pdfDocument, PageSize.A4.Rotate());
                    doc.SetMargins(75, 20, 5, 20);
                }


                //string pathLogo = System.Web.HttpContext.Current.Server.MapPath("~/Images/logoDolarTec.png");

                //Image logotipo = new Image(ImageDataFactory.Create(logo));
                string imgBase64 = "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";
                byte[] imgBytes = Convert.FromBase64String(imgBase64);
                ImageData id = ImageDataFactory.Create(imgBytes);
                Image logotipo = new Image(id);
                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new HeaderEventHandler1(logotipo));
                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new FooterEventHandler1(usuario));

                int nTablas = dst.Tables.Count;

                for (int i = 0; i < nTablas; i++)
                {
                    doc.Add(serializarTablaPDF(doc, dst.Tables[i], Titulo, dst.Tables[i].Prefix != ""));
                    doc.Add(new Paragraph(" "));
                }
                //if (!String.IsNullOrEmpty(imagenBase64))
                //{
                //    doc.Add(new Paragraph(" "));
                //    byte[] bufferImagen = Convert.FromBase64String(imagenBase64);
                //    using (MemoryStream msImagen = new MemoryStream(bufferImagen))
                //    {
                //        System.Drawing.Image imagen = System.Drawing.Image.FromStream(msImagen);
                //        System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(imagen, new System.Drawing.Size(imagenAncho, imagenAlto));
                //      Image pic = new Image(ImageDataFactory.Create(imagen));
                //        doc.Add(pic);
                //    }
                //}
                doc.Close();
                
                buffer = ms.ToArray();
            }
            return buffer;
        }

        public class HeaderEventHandler1 : IEventHandler
        {
            Image Img;
            public HeaderEventHandler1(Image img)
            {
                Img = img;
            }
            public void HandleEvent(Event @event)
            {
                PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
                PdfDocument pdfDoc = docEvent.GetDocument();
                PdfPage page = docEvent.GetPage();

                PdfCanvas canvas1 = new PdfCanvas(page.NewContentStreamBefore(), page.GetResources(), pdfDoc);
                Rectangle rootArea = new Rectangle(35, page.GetPageSize().GetTop() - 75, page.GetPageSize().GetWidth() - 70, 55);
                new Canvas(canvas1, pdfDoc, rootArea).Add(getTable(docEvent));
            }

            public Table getTable(PdfDocumentEvent docEvent)
            {
                float[] cellWith = { 20f, 80f };
                Table tableEvent = new Table(UnitValue.CreatePercentArray(cellWith)).UseAllAvailableWidth();

                Style styleCell = new Style()
                    .SetBorder(Border.NO_BORDER);

                Style styleText = new Style()
                    .SetTextAlignment(TextAlignment.RIGHT).SetFontSize(10f)
                    .SetBold();

                Cell cell = new Cell().Add(Img.SetAutoScale(true));

                tableEvent.AddCell(cell
                    .AddStyle(styleCell)
                    .SetTextAlignment(TextAlignment.LEFT));

                cell = new Cell()
                    .Add(new Paragraph("Estudio Muniz\n"))
                    .Add(new Paragraph("Fecha : " + DateTime.Now.ToShortDateString() + "\n"))
                    .Add(new Paragraph("Hora : " + DateTime.Now.ToShortTimeString()))
                    .AddStyle(styleText).AddStyle(styleCell);

                tableEvent.AddCell(cell);

                return tableEvent;
            }
        }

        public class FooterEventHandler1 : IEventHandler
        {
            string Usuario;
            public FooterEventHandler1(string usuario)
            {
                Usuario = usuario;
            }

            public void HandleEvent(Event @event)
            {
                PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
                PdfDocument pdfDoc = docEvent.GetDocument();
                PdfPage page = docEvent.GetPage();

                PdfCanvas canvas = new PdfCanvas(page.NewContentStreamBefore(), page.GetResources(), pdfDoc);
                Rectangle rootArea = new Rectangle(36, 20, page.GetPageSize().GetWidth() - 70, 55);
                new Canvas(canvas, pdfDoc, rootArea)
                    .Add(getTable(docEvent));
            }

            public Table getTable(PdfDocumentEvent docEvent)
            {
                float[] cellWith = { 92f, 8f };
                Table tableEvent = new Table(UnitValue.CreatePercentArray(cellWith)).UseAllAvailableWidth();

                PdfPage page = docEvent.GetPage();
                int pageNum = docEvent.GetDocument().GetPageNumber(page);

                //int paginasNum = docEvent.GetDocument().GetNumberOfPages();

                Style styleCell = new Style()
                    .SetPadding(5)
                    .SetBorder(Border.NO_BORDER)
                    .SetBorderTop(new SolidBorder(ColorConstants.BLACK, 2));

                Cell cell = new Cell().Add(new Paragraph(Usuario));

                tableEvent.AddCell(cell
                    .AddStyle(styleCell)
                    .SetTextAlignment(TextAlignment.LEFT)
                    .SetBold());

                cell = new Cell().Add(new Paragraph(pageNum.ToString()));
                cell.AddStyle(styleCell)
                    .SetBackgroundColor(ColorConstants.BLACK)
                    .SetFontColor(ColorConstants.WHITE)
                    .SetTextAlignment(TextAlignment.CENTER);
                tableEvent.AddCell(cell);
                return tableEvent;
            }
        }
    }
}