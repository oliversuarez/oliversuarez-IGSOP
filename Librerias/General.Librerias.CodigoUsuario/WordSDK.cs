using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using M = DocumentFormat.OpenXml.Math;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using V = DocumentFormat.OpenXml.Vml;
using Thm15 = DocumentFormat.OpenXml.Office2013.Theme;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace General.Librerias.CodigoUsuario
{
    public class WordSDK
    {
        string titulo;
        string subtituloIzquierda;
        string subtituloDerecha;
        DataTable tabla;
        Dictionary<string, string> piePagina = null;
        int nCols;
        int nFilas;
        bool horizontal = false;

        public void CrearArchivo(string archivoExcel, string _titulo, string _subtituloIzquierda, string _subtituloDerecha, DataTable _tabla, Dictionary<string, string> _piePagina = null, bool _horizontal = false)
        {
            titulo = _titulo;
            subtituloIzquierda = _subtituloIzquierda;
            subtituloDerecha = _subtituloDerecha;
            tabla = _tabla;
            piePagina = _piePagina;
            horizontal = _horizontal;
            nCols = tabla.Columns.Count;
            nFilas = tabla.Rows.Count;
            using (WordprocessingDocument package = WordprocessingDocument.Create(archivoExcel, WordprocessingDocumentType.Document))
            {
                CreateParts(package);
            }
        }

        public byte[] CrearBytes(string _titulo, string _subtituloIzquierda, string _subtituloDerecha, DataTable _tabla, Dictionary<string, string> _piePagina = null, bool _horizontal = false)
        {
            byte[] rpta = null;
            titulo = _titulo;
            subtituloIzquierda = _subtituloIzquierda;
            subtituloDerecha = _subtituloDerecha;
            tabla = _tabla;
            piePagina = _piePagina;
            horizontal = _horizontal;
            nCols = tabla.Columns.Count;
            nFilas = tabla.Rows.Count;
            using (MemoryStream ms = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
                {
                    CreateParts(package);
                }
                rpta = ms.ToArray();
            }
            return rpta;
        }

        private void CreateParts(WordprocessingDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            MainDocumentPart mainDocumentPart1 = document.AddMainDocumentPart();
            GenerateMainDocumentPart1Content(mainDocumentPart1);

            WebSettingsPart webSettingsPart1 = mainDocumentPart1.AddNewPart<WebSettingsPart>("rId3");
            GenerateWebSettingsPart1Content(webSettingsPart1);

            DocumentSettingsPart documentSettingsPart1 = mainDocumentPart1.AddNewPart<DocumentSettingsPart>("rId2");
            GenerateDocumentSettingsPart1Content(documentSettingsPart1);

            StyleDefinitionsPart styleDefinitionsPart1 = mainDocumentPart1.AddNewPart<StyleDefinitionsPart>("rId1");
            GenerateStyleDefinitionsPart1Content(styleDefinitionsPart1);

            ThemePart themePart1 = mainDocumentPart1.AddNewPart<ThemePart>("rId6");
            GenerateThemePart1Content(themePart1);

            FontTablePart fontTablePart1 = mainDocumentPart1.AddNewPart<FontTablePart>("rId5");
            GenerateFontTablePart1Content(fontTablePart1);

            ImagePart imagePart1 = mainDocumentPart1.AddNewPart<ImagePart>("image/jpeg", "rId4");
            GenerateImagePart1Content(imagePart1);

            SetPackageProperties(document);
        }

        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Template template1 = new Ap.Template();
            template1.Text = "Normal.dotm";
            Ap.TotalTime totalTime1 = new Ap.TotalTime();
            totalTime1.Text = "6";
            Ap.Pages pages1 = new Ap.Pages();
            pages1.Text = "1";
            Ap.Words words1 = new Ap.Words();
            words1.Text = "22";
            Ap.Characters characters1 = new Ap.Characters();
            characters1.Text = "127";
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Office Word";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.Lines lines1 = new Ap.Lines();
            lines1.Text = "1";
            Ap.Paragraphs paragraphs1 = new Ap.Paragraphs();
            paragraphs1.Text = "1";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.CharactersWithSpaces charactersWithSpaces1 = new Ap.CharactersWithSpaces();
            charactersWithSpaces1.Text = "148";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(template1);
            properties1.Append(totalTime1);
            properties1.Append(pages1);
            properties1.Append(words1);
            properties1.Append(characters1);
            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(lines1);
            properties1.Append(paragraphs1);
            properties1.Append(scaleCrop1);
            properties1.Append(company1);
            properties1.Append(linksUpToDate1);
            properties1.Append(charactersWithSpaces1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void GenerateMainDocumentPart1Content(MainDocumentPart mainDocumentPart1)
        {
            int ancho;
            int anchoTotal = 0;
            int escala = 27;
            int nMitad = (int)(nCols / 2);
            int anchoMitad = 0;

            Document document1 = new Document();
            document1.AddNamespaceDeclaration("ve", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            document1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            document1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            document1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            document1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            document1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            document1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            document1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            document1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");

            Body body1 = new Body();

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "Tablaconcuadrcula" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);
            table1.Append(tableProperties1);            

            TableGrid tableGrid1 = new TableGrid();
            for(int j = 0; j < nCols; j++)
            {
                ancho = int.Parse(tabla.Columns[j].Caption);
                anchoTotal += ancho;
                if (j < nMitad) anchoMitad += ancho;
                GridColumn gridColumn1 = new GridColumn() { Width = (ancho*escala).ToString() };
                tableGrid1.Append(gridColumn1);
            }
            table1.Append(tableGrid1);

            //Crear la Primera Fila con la Imagen del Logo
            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "00E553F0", RsidTableRowProperties = "001C4887" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = (anchoTotal*escala).ToString(), Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan1 = new GridSpan() { Val = nCols };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(gridSpan1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00E553F0", RsidRunAdditionDefault = "00E553F0" };

            Run run1 = new Run() { RsidRunProperties = "00E553F0" };

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U };
            Wp.Extent extent1 = new Wp.Extent() { Cx = 1891665L, Cy = 647700L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };

            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Imagen 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();
            nonVisualDrawingPropertiesExtensionList1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" xmlns:lc=\"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas\" id=\"{CB2255AE-3FF9-453F-B440-93DF0C0596EF}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            docProperties1.Append(nonVisualDrawingPropertiesExtensionList1);
            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();

            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList2 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension2 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" xmlns:lc=\"http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas\" id=\"{CB2255AE-3FF9-453F-B440-93DF0C0596EF}\" />");

            nonVisualDrawingPropertiesExtension2.Append(openXmlUnknownElement2);

            nonVisualDrawingPropertiesExtensionList2.Append(nonVisualDrawingPropertiesExtension2);

            nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList2);

            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId4", CompressionState = A.BlipCompressionValues.Print };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            useLocalDpi1.AddNamespaceDeclaration("lc", "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);
            A.SourceRectangle sourceRectangle1 = new A.SourceRectangle();

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(sourceRectangle1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 1891665L, Cy = 647700L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            A.ShapePropertiesExtensionList shapePropertiesExtensionList1 = new A.ShapePropertiesExtensionList();

            A.ShapePropertiesExtension shapePropertiesExtension1 = new A.ShapePropertiesExtension() { Uri = "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}" };

            A14.HiddenFillProperties hiddenFillProperties1 = new A14.HiddenFillProperties();
            hiddenFillProperties1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
            hiddenFillProperties1.AddNamespaceDeclaration("lc", "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas");

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill1.Append(rgbColorModelHex1);

            hiddenFillProperties1.Append(solidFill1);

            shapePropertiesExtension1.Append(hiddenFillProperties1);

            shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(shapePropertiesExtensionList1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(drawing1);

            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            tableRow1.Append(tableCell1);
            table1.Append(tableRow1);

            //Crear la Segunda Fila con el Titulo
            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "00E553F0", RsidTableRowProperties = "00E53055" };

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = (anchoTotal * escala).ToString(), Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan2 = new GridSpan() { Val = nCols };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(gridSpan2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "003A6C84", RsidParagraphAddition = "00E553F0", RsidParagraphProperties = "003A6C84", RsidRunAdditionDefault = "003A6C84" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Center };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Bold bold1 = new Bold();

            paragraphMarkRunProperties1.Append(bold1);

            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run2 = new Run() { RsidRunProperties = "003A6C84" };

            RunProperties runProperties1 = new RunProperties();
            Bold bold2 = new Bold();
            FontSize fontSize1 = new FontSize() { Val = "36" };

            runProperties1.Append(bold2);
            runProperties1.Append(fontSize1);
            Text text1 = new Text();
            text1.Text = titulo;

            run2.Append(runProperties1);
            run2.Append(text1);

            paragraph2.Append(paragraphProperties1);
            paragraph2.Append(run2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            tableRow2.Append(tableCell2);
            table1.Append(tableRow2);

            //Crear la Tercera Fila con las celdas de subtitulos
            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "00E553F0", RsidTableRowProperties = "00906834" };

            //Crear la Celda con el Subtitulo de la Izquierda
            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = (anchoMitad * escala).ToString(), Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan3 = new GridSpan() { Val = nMitad };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(gridSpan3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "003A6C84", RsidParagraphAddition = "00E553F0", RsidRunAdditionDefault = "003A6C84" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            Bold bold3 = new Bold();

            paragraphMarkRunProperties2.Append(bold3);

            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run3 = new Run() { RsidRunProperties = "003A6C84" };

            RunProperties runProperties2 = new RunProperties();
            Bold bold4 = new Bold();

            runProperties2.Append(bold4);
            Text text2 = new Text();
            text2.Text = subtituloIzquierda;

            run3.Append(runProperties2);
            run3.Append(text2);

            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(run3);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            //Crear la Celda con el Subtitulo de la Derecha
            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = ((anchoTotal - anchoMitad) * escala).ToString(), Type = TableWidthUnitValues.Dxa };
            GridSpan gridSpan4 = new GridSpan() { Val = (nCols - nMitad) };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(gridSpan4);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "003A6C84", RsidParagraphAddition = "00E553F0", RsidParagraphProperties = "003A6C84", RsidRunAdditionDefault = "003A6C84" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Right };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            Bold bold5 = new Bold();

            paragraphMarkRunProperties3.Append(bold5);

            paragraphProperties3.Append(justification2);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run4 = new Run() { RsidRunProperties = "003A6C84" };

            RunProperties runProperties3 = new RunProperties();
            Bold bold6 = new Bold();

            runProperties3.Append(bold6);
            Text text3 = new Text();
            text3.Text = subtituloDerecha;

            run4.Append(runProperties3);
            run4.Append(text3);

            paragraph4.Append(paragraphProperties3);
            paragraph4.Append(run4);

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            tableRow3.Append(tableCell3);
            tableRow3.Append(tableCell4);
            table1.Append(tableRow3);            

            //Crear la Fila de las Cabaceras de la Data
            TableRow tableRow4 = new TableRow() { RsidTableRowAddition = "00E553F0", RsidTableRowProperties = "00E553F0" };
            for(int j = 0; j < nCols; j++)
            {
                ancho = int.Parse(tabla.Columns[j].Caption);

                TableCell tableCell5 = new TableCell();

                TableCellProperties tableCellProperties5 = new TableCellProperties();
                TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = (escala * ancho).ToString(), Type = TableWidthUnitValues.Dxa };

                tableCellProperties5.Append(tableCellWidth5);

                Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "003A6C84", RsidParagraphAddition = "00E553F0", RsidRunAdditionDefault = "00E553F0" };

                ParagraphProperties paragraphProperties4 = new ParagraphProperties();

                ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                Bold bold7 = new Bold();
                Color color1 = new Color() { Val = "92D050" };

                paragraphMarkRunProperties4.Append(bold7);
                paragraphMarkRunProperties4.Append(color1);

                paragraphProperties4.Append(paragraphMarkRunProperties4);

                Run run5 = new Run() { RsidRunProperties = "003A6C84" };

                RunProperties runProperties4 = new RunProperties();
                Bold bold8 = new Bold();
                Color color2 = new Color() { Val = "92D050" };

                runProperties4.Append(bold8);
                runProperties4.Append(color2);
                Text text4 = new Text();
                text4.Text = tabla.Columns[j].ColumnName;

                run5.Append(runProperties4);
                run5.Append(text4);

                paragraph5.Append(paragraphProperties4);
                paragraph5.Append(run5);

                tableCell5.Append(tableCellProperties5);
                tableCell5.Append(paragraph5);

                tableRow4.Append(tableCell5);
            }
            table1.Append(tableRow4);

            //Crear las Filas con la Data
            for(int i = 0; i < nFilas; i++)
            {
                TableRow tableRow5 = new TableRow() { RsidTableRowAddition = "00E553F0", RsidTableRowProperties = "00E553F0" };
                for(int j = 0; j < nCols; j++)
                {
                    ancho = int.Parse(tabla.Columns[j].Caption);
                    TableCell tableCell6 = new TableCell();

                    TableCellProperties tableCellProperties6 = new TableCellProperties();
                    TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "1080", Type = TableWidthUnitValues.Dxa };

                    tableCellProperties6.Append(tableCellWidth6);

                    Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00E553F0", RsidRunAdditionDefault = "00E553F0" };

                    Run run6 = new Run();
                    Text text6 = new Text();
                    text6.Text = tabla.Rows[i][j].ToString();

                    run6.Append(text6);

                    paragraph6.Append(run6);

                    tableCell6.Append(tableCellProperties6);
                    tableCell6.Append(paragraph6);

                    tableRow5.Append(tableCell6);
                }
                table1.Append(tableRow5);
            }

            if(piePagina!=null && piePagina.Count > 0)
            {
                //Crear la Fila de las Cabaceras de la Data
                TableRow tableRow6 = new TableRow() { RsidTableRowAddition = "00E553F0", RsidTableRowProperties = "00E553F0" };
                for (int j = 0; j < nCols; j++)
                {                    
                    ancho = int.Parse(tabla.Columns[j].Caption);

                    TableCell tableCell7 = new TableCell();

                    TableCellProperties tableCellProperties7 = new TableCellProperties();
                    TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = (escala * ancho).ToString(), Type = TableWidthUnitValues.Dxa };

                    tableCellProperties7.Append(tableCellWidth7);

                    Paragraph paragraph7 = new Paragraph() { RsidParagraphMarkRevision = "003A6C84", RsidParagraphAddition = "00E553F0", RsidRunAdditionDefault = "00E553F0" };

                    ParagraphProperties paragraphProperties7 = new ParagraphProperties();

                    ParagraphMarkRunProperties paragraphMarkRunProperties4 = new ParagraphMarkRunProperties();
                    Bold bold7 = new Bold();
                    //Color color1 = new Color() { Val = "92D050" };

                    paragraphMarkRunProperties4.Append(bold7);
                    //paragraphMarkRunProperties4.Append(color1);

                    paragraphProperties7.Append(paragraphMarkRunProperties4);

                    Run run7 = new Run() { RsidRunProperties = "003A6C84" };

                    RunProperties runProperties7 = new RunProperties();
                    Bold bold8 = new Bold();
                    //Color color8 = new Color() { Val = "92D050" };
                    runProperties7.Append(bold8);
                    //runProperties7.Append(color8);
                    run7.Append(runProperties7);

                    if (piePagina.ContainsKey(j.ToString()))
                    {
                        Text text4 = new Text();
                        text4.Text = piePagina[j.ToString()];
                        run7.Append(text4);
                    }

                    paragraph7.Append(paragraphProperties7);
                    paragraph7.Append(run7);

                    tableCell7.Append(paragraph7);
                    tableCell7.Append(tableCellProperties7);
                    tableRow6.Append(tableCell7);                 
                }
                table1.Append(tableRow6);
            }

            Paragraph paragraph21 = new Paragraph() { RsidParagraphAddition = "0085037B", RsidRunAdditionDefault = "0085037B" };

            SectionProperties sectionProperties1 = new SectionProperties() { RsidR = "0085037B", RsidSect = "003D3FA0" };
            UInt32Value anchoPagina = (UInt32Value)11906U;
            UInt32Value altoPagina = (UInt32Value)16838U;
            PageOrientationValues orientacion = PageOrientationValues.Portrait;
            if (horizontal)
            {
                anchoPagina = (UInt32Value)16838U;
                altoPagina = (UInt32Value)11906U;
                orientacion = PageOrientationValues.Landscape;
            }
            PageSize pageSize1 = new PageSize() { Width = anchoPagina, Height = altoPagina, Orient = orientacion };
            
            PageMargin pageMargin1 = new PageMargin() { Top = 1417, Right = (UInt32Value)1701U, Bottom = 1417, Left = (UInt32Value)1701U, Header = (UInt32Value)708U, Footer = (UInt32Value)708U, Gutter = (UInt32Value)0U };
            Columns columns1 = new Columns() { Space = "708" };
            DocGrid docGrid1 = new DocGrid() { LinePitch = 360 };

            sectionProperties1.Append(pageSize1);
            sectionProperties1.Append(pageMargin1);
            sectionProperties1.Append(columns1);
            sectionProperties1.Append(docGrid1);

            body1.Append(table1);
            //body1.Append(paragraph21);
            body1.Append(sectionProperties1);
            document1.Append(body1);
            mainDocumentPart1.Document = document1;
        }

        private void GenerateWebSettingsPart1Content(WebSettingsPart webSettingsPart1)
        {
            WebSettings webSettings1 = new WebSettings();
            webSettings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            webSettings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            OptimizeForBrowser optimizeForBrowser1 = new OptimizeForBrowser();

            webSettings1.Append(optimizeForBrowser1);

            webSettingsPart1.WebSettings = webSettings1;
        }

        private void GenerateDocumentSettingsPart1Content(DocumentSettingsPart documentSettingsPart1)
        {
            Settings settings1 = new Settings();
            settings1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            settings1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            settings1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            settings1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            settings1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            settings1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            settings1.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            Zoom zoom1 = new Zoom() { Percent = "100" };
            ProofState proofState1 = new ProofState() { Spelling = ProofingStateValues.Clean, Grammar = ProofingStateValues.Clean };
            DefaultTabStop defaultTabStop1 = new DefaultTabStop() { Val = 708 };
            HyphenationZone hyphenationZone1 = new HyphenationZone() { Val = "425" };
            CharacterSpacingControl characterSpacingControl1 = new CharacterSpacingControl() { Val = CharacterSpacingValues.DoNotCompress };
            Compatibility compatibility1 = new Compatibility();

            Rsids rsids1 = new Rsids();
            RsidRoot rsidRoot1 = new RsidRoot() { Val = "00E553F0" };
            Rsid rsid1 = new Rsid() { Val = "00027849" };
            Rsid rsid2 = new Rsid() { Val = "00355FE2" };
            Rsid rsid3 = new Rsid() { Val = "003A6C84" };
            Rsid rsid4 = new Rsid() { Val = "003D3FA0" };
            Rsid rsid5 = new Rsid() { Val = "00435B6B" };
            Rsid rsid6 = new Rsid() { Val = "0085037B" };
            Rsid rsid7 = new Rsid() { Val = "00854D17" };
            Rsid rsid8 = new Rsid() { Val = "00AC0ECB" };
            Rsid rsid9 = new Rsid() { Val = "00E553F0" };
            Rsid rsid10 = new Rsid() { Val = "00E962BE" };
            Rsid rsid11 = new Rsid() { Val = "00EC1EF3" };

            rsids1.Append(rsidRoot1);
            rsids1.Append(rsid1);
            rsids1.Append(rsid2);
            rsids1.Append(rsid3);
            rsids1.Append(rsid4);
            rsids1.Append(rsid5);
            rsids1.Append(rsid6);
            rsids1.Append(rsid7);
            rsids1.Append(rsid8);
            rsids1.Append(rsid9);
            rsids1.Append(rsid10);
            rsids1.Append(rsid11);

            M.MathProperties mathProperties1 = new M.MathProperties();
            M.MathFont mathFont1 = new M.MathFont() { Val = "Cambria Math" };
            M.BreakBinary breakBinary1 = new M.BreakBinary() { Val = M.BreakBinaryOperatorValues.Before };
            M.BreakBinarySubtraction breakBinarySubtraction1 = new M.BreakBinarySubtraction() { Val = M.BreakBinarySubtractionValues.MinusMinus };
            M.SmallFraction smallFraction1 = new M.SmallFraction();
            M.DisplayDefaults displayDefaults1 = new M.DisplayDefaults();
            M.LeftMargin leftMargin1 = new M.LeftMargin() { Val = (UInt32Value)0U };
            M.RightMargin rightMargin1 = new M.RightMargin() { Val = (UInt32Value)0U };
            M.DefaultJustification defaultJustification1 = new M.DefaultJustification() { Val = M.JustificationValues.CenterGroup };
            M.WrapIndent wrapIndent1 = new M.WrapIndent() { Val = (UInt32Value)1440U };
            M.IntegralLimitLocation integralLimitLocation1 = new M.IntegralLimitLocation() { Val = M.LimitLocationValues.SubscriptSuperscript };
            M.NaryLimitLocation naryLimitLocation1 = new M.NaryLimitLocation() { Val = M.LimitLocationValues.UnderOver };

            mathProperties1.Append(mathFont1);
            mathProperties1.Append(breakBinary1);
            mathProperties1.Append(breakBinarySubtraction1);
            mathProperties1.Append(smallFraction1);
            mathProperties1.Append(displayDefaults1);
            mathProperties1.Append(leftMargin1);
            mathProperties1.Append(rightMargin1);
            mathProperties1.Append(defaultJustification1);
            mathProperties1.Append(wrapIndent1);
            mathProperties1.Append(integralLimitLocation1);
            mathProperties1.Append(naryLimitLocation1);
            ThemeFontLanguages themeFontLanguages1 = new ThemeFontLanguages() { Val = "es-PE" };
            ColorSchemeMapping colorSchemeMapping1 = new ColorSchemeMapping() { Background1 = ColorSchemeIndexValues.Light1, Text1 = ColorSchemeIndexValues.Dark1, Background2 = ColorSchemeIndexValues.Light2, Text2 = ColorSchemeIndexValues.Dark2, Accent1 = ColorSchemeIndexValues.Accent1, Accent2 = ColorSchemeIndexValues.Accent2, Accent3 = ColorSchemeIndexValues.Accent3, Accent4 = ColorSchemeIndexValues.Accent4, Accent5 = ColorSchemeIndexValues.Accent5, Accent6 = ColorSchemeIndexValues.Accent6, Hyperlink = ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = ColorSchemeIndexValues.FollowedHyperlink };

            ShapeDefaults shapeDefaults1 = new ShapeDefaults();
            Ovml.ShapeDefaults shapeDefaults2 = new Ovml.ShapeDefaults() { Extension = V.ExtensionHandlingBehaviorValues.Edit, MaxShapeId = 2050 };

            Ovml.ShapeLayout shapeLayout1 = new Ovml.ShapeLayout() { Extension = V.ExtensionHandlingBehaviorValues.Edit };
            Ovml.ShapeIdMap shapeIdMap1 = new Ovml.ShapeIdMap() { Extension = V.ExtensionHandlingBehaviorValues.Edit, Data = "1" };

            shapeLayout1.Append(shapeIdMap1);

            shapeDefaults1.Append(shapeDefaults2);
            shapeDefaults1.Append(shapeLayout1);
            DecimalSymbol decimalSymbol1 = new DecimalSymbol() { Val = "." };
            ListSeparator listSeparator1 = new ListSeparator() { Val = "," };

            settings1.Append(zoom1);
            settings1.Append(proofState1);
            settings1.Append(defaultTabStop1);
            settings1.Append(hyphenationZone1);
            settings1.Append(characterSpacingControl1);
            settings1.Append(compatibility1);
            settings1.Append(rsids1);
            settings1.Append(mathProperties1);
            settings1.Append(themeFontLanguages1);
            settings1.Append(colorSchemeMapping1);
            settings1.Append(shapeDefaults1);
            settings1.Append(decimalSymbol1);
            settings1.Append(listSeparator1);

            documentSettingsPart1.Settings = settings1;
        }

        private void GenerateStyleDefinitionsPart1Content(StyleDefinitionsPart styleDefinitionsPart1)
        {
            Styles styles1 = new Styles();
            styles1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            styles1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            DocDefaults docDefaults1 = new DocDefaults();

            RunPropertiesDefault runPropertiesDefault1 = new RunPropertiesDefault();

            RunPropertiesBaseStyle runPropertiesBaseStyle1 = new RunPropertiesBaseStyle();
            RunFonts runFonts1 = new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi };
            FontSize fontSize2 = new FontSize() { Val = "22" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "22" };
            Languages languages1 = new Languages() { Val = "es-PE", EastAsia = "en-US", Bidi = "ar-SA" };

            runPropertiesBaseStyle1.Append(runFonts1);
            runPropertiesBaseStyle1.Append(fontSize2);
            runPropertiesBaseStyle1.Append(fontSizeComplexScript1);
            runPropertiesBaseStyle1.Append(languages1);

            runPropertiesDefault1.Append(runPropertiesBaseStyle1);

            ParagraphPropertiesDefault paragraphPropertiesDefault1 = new ParagraphPropertiesDefault();

            ParagraphPropertiesBaseStyle paragraphPropertiesBaseStyle1 = new ParagraphPropertiesBaseStyle();
            SpacingBetweenLines spacingBetweenLines1 = new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto };

            paragraphPropertiesBaseStyle1.Append(spacingBetweenLines1);

            paragraphPropertiesDefault1.Append(paragraphPropertiesBaseStyle1);

            docDefaults1.Append(runPropertiesDefault1);
            docDefaults1.Append(paragraphPropertiesDefault1);

            LatentStyles latentStyles1 = new LatentStyles() { DefaultLockedState = false, DefaultUiPriority = 99, DefaultSemiHidden = true, DefaultUnhideWhenUsed = true, DefaultPrimaryStyle = false, Count = 267 };
            LatentStyleExceptionInfo latentStyleExceptionInfo1 = new LatentStyleExceptionInfo() { Name = "Normal", UiPriority = 0, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo2 = new LatentStyleExceptionInfo() { Name = "heading 1", UiPriority = 9, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo3 = new LatentStyleExceptionInfo() { Name = "heading 2", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo4 = new LatentStyleExceptionInfo() { Name = "heading 3", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo5 = new LatentStyleExceptionInfo() { Name = "heading 4", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo6 = new LatentStyleExceptionInfo() { Name = "heading 5", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo7 = new LatentStyleExceptionInfo() { Name = "heading 6", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo8 = new LatentStyleExceptionInfo() { Name = "heading 7", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo9 = new LatentStyleExceptionInfo() { Name = "heading 8", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo10 = new LatentStyleExceptionInfo() { Name = "heading 9", UiPriority = 9, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo11 = new LatentStyleExceptionInfo() { Name = "toc 1", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo12 = new LatentStyleExceptionInfo() { Name = "toc 2", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo13 = new LatentStyleExceptionInfo() { Name = "toc 3", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo14 = new LatentStyleExceptionInfo() { Name = "toc 4", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo15 = new LatentStyleExceptionInfo() { Name = "toc 5", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo16 = new LatentStyleExceptionInfo() { Name = "toc 6", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo17 = new LatentStyleExceptionInfo() { Name = "toc 7", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo18 = new LatentStyleExceptionInfo() { Name = "toc 8", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo19 = new LatentStyleExceptionInfo() { Name = "toc 9", UiPriority = 39 };
            LatentStyleExceptionInfo latentStyleExceptionInfo20 = new LatentStyleExceptionInfo() { Name = "caption", UiPriority = 35, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo21 = new LatentStyleExceptionInfo() { Name = "Title", UiPriority = 10, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo22 = new LatentStyleExceptionInfo() { Name = "Default Paragraph Font", UiPriority = 1 };
            LatentStyleExceptionInfo latentStyleExceptionInfo23 = new LatentStyleExceptionInfo() { Name = "Subtitle", UiPriority = 11, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo24 = new LatentStyleExceptionInfo() { Name = "Strong", UiPriority = 22, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo25 = new LatentStyleExceptionInfo() { Name = "Emphasis", UiPriority = 20, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo26 = new LatentStyleExceptionInfo() { Name = "Table Grid", UiPriority = 39, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo27 = new LatentStyleExceptionInfo() { Name = "Placeholder Text", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo28 = new LatentStyleExceptionInfo() { Name = "No Spacing", UiPriority = 1, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo29 = new LatentStyleExceptionInfo() { Name = "Light Shading", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo30 = new LatentStyleExceptionInfo() { Name = "Light List", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo31 = new LatentStyleExceptionInfo() { Name = "Light Grid", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo32 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo33 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo34 = new LatentStyleExceptionInfo() { Name = "Medium List 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo35 = new LatentStyleExceptionInfo() { Name = "Medium List 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo36 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo37 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo38 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo39 = new LatentStyleExceptionInfo() { Name = "Dark List", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo40 = new LatentStyleExceptionInfo() { Name = "Colorful Shading", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo41 = new LatentStyleExceptionInfo() { Name = "Colorful List", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo42 = new LatentStyleExceptionInfo() { Name = "Colorful Grid", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo43 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 1", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo44 = new LatentStyleExceptionInfo() { Name = "Light List Accent 1", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo45 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 1", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo46 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 1", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo47 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 1", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo48 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 1", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo49 = new LatentStyleExceptionInfo() { Name = "Revision", UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo50 = new LatentStyleExceptionInfo() { Name = "List Paragraph", UiPriority = 34, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo51 = new LatentStyleExceptionInfo() { Name = "Quote", UiPriority = 29, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo52 = new LatentStyleExceptionInfo() { Name = "Intense Quote", UiPriority = 30, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo53 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 1", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo54 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 1", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo55 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 1", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo56 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 1", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo57 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 1", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo58 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 1", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo59 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 1", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo60 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 1", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo61 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 2", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo62 = new LatentStyleExceptionInfo() { Name = "Light List Accent 2", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo63 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 2", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo64 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 2", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo65 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 2", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo66 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 2", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo67 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 2", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo68 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 2", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo69 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 2", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo70 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 2", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo71 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 2", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo72 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 2", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo73 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 2", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo74 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 2", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo75 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 3", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo76 = new LatentStyleExceptionInfo() { Name = "Light List Accent 3", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo77 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 3", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo78 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 3", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo79 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 3", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo80 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 3", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo81 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 3", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo82 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 3", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo83 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 3", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo84 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 3", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo85 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 3", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo86 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 3", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo87 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 3", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo88 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 3", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo89 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 4", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo90 = new LatentStyleExceptionInfo() { Name = "Light List Accent 4", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo91 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 4", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo92 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 4", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo93 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 4", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo94 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 4", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo95 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 4", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo96 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 4", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo97 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 4", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo98 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 4", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo99 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 4", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo100 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 4", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo101 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 4", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo102 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 4", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo103 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 5", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo104 = new LatentStyleExceptionInfo() { Name = "Light List Accent 5", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo105 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 5", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo106 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 5", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo107 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 5", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo108 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 5", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo109 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 5", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo110 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 5", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo111 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 5", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo112 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 5", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo113 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 5", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo114 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 5", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo115 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 5", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo116 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 5", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo117 = new LatentStyleExceptionInfo() { Name = "Light Shading Accent 6", UiPriority = 60, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo118 = new LatentStyleExceptionInfo() { Name = "Light List Accent 6", UiPriority = 61, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo119 = new LatentStyleExceptionInfo() { Name = "Light Grid Accent 6", UiPriority = 62, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo120 = new LatentStyleExceptionInfo() { Name = "Medium Shading 1 Accent 6", UiPriority = 63, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo121 = new LatentStyleExceptionInfo() { Name = "Medium Shading 2 Accent 6", UiPriority = 64, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo122 = new LatentStyleExceptionInfo() { Name = "Medium List 1 Accent 6", UiPriority = 65, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo123 = new LatentStyleExceptionInfo() { Name = "Medium List 2 Accent 6", UiPriority = 66, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo124 = new LatentStyleExceptionInfo() { Name = "Medium Grid 1 Accent 6", UiPriority = 67, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo125 = new LatentStyleExceptionInfo() { Name = "Medium Grid 2 Accent 6", UiPriority = 68, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo126 = new LatentStyleExceptionInfo() { Name = "Medium Grid 3 Accent 6", UiPriority = 69, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo127 = new LatentStyleExceptionInfo() { Name = "Dark List Accent 6", UiPriority = 70, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo128 = new LatentStyleExceptionInfo() { Name = "Colorful Shading Accent 6", UiPriority = 71, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo129 = new LatentStyleExceptionInfo() { Name = "Colorful List Accent 6", UiPriority = 72, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo130 = new LatentStyleExceptionInfo() { Name = "Colorful Grid Accent 6", UiPriority = 73, SemiHidden = false, UnhideWhenUsed = false };
            LatentStyleExceptionInfo latentStyleExceptionInfo131 = new LatentStyleExceptionInfo() { Name = "Subtle Emphasis", UiPriority = 19, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo132 = new LatentStyleExceptionInfo() { Name = "Intense Emphasis", UiPriority = 21, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo133 = new LatentStyleExceptionInfo() { Name = "Subtle Reference", UiPriority = 31, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo134 = new LatentStyleExceptionInfo() { Name = "Intense Reference", UiPriority = 32, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo135 = new LatentStyleExceptionInfo() { Name = "Book Title", UiPriority = 33, SemiHidden = false, UnhideWhenUsed = false, PrimaryStyle = true };
            LatentStyleExceptionInfo latentStyleExceptionInfo136 = new LatentStyleExceptionInfo() { Name = "Bibliography", UiPriority = 37 };
            LatentStyleExceptionInfo latentStyleExceptionInfo137 = new LatentStyleExceptionInfo() { Name = "TOC Heading", UiPriority = 39, PrimaryStyle = true };

            latentStyles1.Append(latentStyleExceptionInfo1);
            latentStyles1.Append(latentStyleExceptionInfo2);
            latentStyles1.Append(latentStyleExceptionInfo3);
            latentStyles1.Append(latentStyleExceptionInfo4);
            latentStyles1.Append(latentStyleExceptionInfo5);
            latentStyles1.Append(latentStyleExceptionInfo6);
            latentStyles1.Append(latentStyleExceptionInfo7);
            latentStyles1.Append(latentStyleExceptionInfo8);
            latentStyles1.Append(latentStyleExceptionInfo9);
            latentStyles1.Append(latentStyleExceptionInfo10);
            latentStyles1.Append(latentStyleExceptionInfo11);
            latentStyles1.Append(latentStyleExceptionInfo12);
            latentStyles1.Append(latentStyleExceptionInfo13);
            latentStyles1.Append(latentStyleExceptionInfo14);
            latentStyles1.Append(latentStyleExceptionInfo15);
            latentStyles1.Append(latentStyleExceptionInfo16);
            latentStyles1.Append(latentStyleExceptionInfo17);
            latentStyles1.Append(latentStyleExceptionInfo18);
            latentStyles1.Append(latentStyleExceptionInfo19);
            latentStyles1.Append(latentStyleExceptionInfo20);
            latentStyles1.Append(latentStyleExceptionInfo21);
            latentStyles1.Append(latentStyleExceptionInfo22);
            latentStyles1.Append(latentStyleExceptionInfo23);
            latentStyles1.Append(latentStyleExceptionInfo24);
            latentStyles1.Append(latentStyleExceptionInfo25);
            latentStyles1.Append(latentStyleExceptionInfo26);
            latentStyles1.Append(latentStyleExceptionInfo27);
            latentStyles1.Append(latentStyleExceptionInfo28);
            latentStyles1.Append(latentStyleExceptionInfo29);
            latentStyles1.Append(latentStyleExceptionInfo30);
            latentStyles1.Append(latentStyleExceptionInfo31);
            latentStyles1.Append(latentStyleExceptionInfo32);
            latentStyles1.Append(latentStyleExceptionInfo33);
            latentStyles1.Append(latentStyleExceptionInfo34);
            latentStyles1.Append(latentStyleExceptionInfo35);
            latentStyles1.Append(latentStyleExceptionInfo36);
            latentStyles1.Append(latentStyleExceptionInfo37);
            latentStyles1.Append(latentStyleExceptionInfo38);
            latentStyles1.Append(latentStyleExceptionInfo39);
            latentStyles1.Append(latentStyleExceptionInfo40);
            latentStyles1.Append(latentStyleExceptionInfo41);
            latentStyles1.Append(latentStyleExceptionInfo42);
            latentStyles1.Append(latentStyleExceptionInfo43);
            latentStyles1.Append(latentStyleExceptionInfo44);
            latentStyles1.Append(latentStyleExceptionInfo45);
            latentStyles1.Append(latentStyleExceptionInfo46);
            latentStyles1.Append(latentStyleExceptionInfo47);
            latentStyles1.Append(latentStyleExceptionInfo48);
            latentStyles1.Append(latentStyleExceptionInfo49);
            latentStyles1.Append(latentStyleExceptionInfo50);
            latentStyles1.Append(latentStyleExceptionInfo51);
            latentStyles1.Append(latentStyleExceptionInfo52);
            latentStyles1.Append(latentStyleExceptionInfo53);
            latentStyles1.Append(latentStyleExceptionInfo54);
            latentStyles1.Append(latentStyleExceptionInfo55);
            latentStyles1.Append(latentStyleExceptionInfo56);
            latentStyles1.Append(latentStyleExceptionInfo57);
            latentStyles1.Append(latentStyleExceptionInfo58);
            latentStyles1.Append(latentStyleExceptionInfo59);
            latentStyles1.Append(latentStyleExceptionInfo60);
            latentStyles1.Append(latentStyleExceptionInfo61);
            latentStyles1.Append(latentStyleExceptionInfo62);
            latentStyles1.Append(latentStyleExceptionInfo63);
            latentStyles1.Append(latentStyleExceptionInfo64);
            latentStyles1.Append(latentStyleExceptionInfo65);
            latentStyles1.Append(latentStyleExceptionInfo66);
            latentStyles1.Append(latentStyleExceptionInfo67);
            latentStyles1.Append(latentStyleExceptionInfo68);
            latentStyles1.Append(latentStyleExceptionInfo69);
            latentStyles1.Append(latentStyleExceptionInfo70);
            latentStyles1.Append(latentStyleExceptionInfo71);
            latentStyles1.Append(latentStyleExceptionInfo72);
            latentStyles1.Append(latentStyleExceptionInfo73);
            latentStyles1.Append(latentStyleExceptionInfo74);
            latentStyles1.Append(latentStyleExceptionInfo75);
            latentStyles1.Append(latentStyleExceptionInfo76);
            latentStyles1.Append(latentStyleExceptionInfo77);
            latentStyles1.Append(latentStyleExceptionInfo78);
            latentStyles1.Append(latentStyleExceptionInfo79);
            latentStyles1.Append(latentStyleExceptionInfo80);
            latentStyles1.Append(latentStyleExceptionInfo81);
            latentStyles1.Append(latentStyleExceptionInfo82);
            latentStyles1.Append(latentStyleExceptionInfo83);
            latentStyles1.Append(latentStyleExceptionInfo84);
            latentStyles1.Append(latentStyleExceptionInfo85);
            latentStyles1.Append(latentStyleExceptionInfo86);
            latentStyles1.Append(latentStyleExceptionInfo87);
            latentStyles1.Append(latentStyleExceptionInfo88);
            latentStyles1.Append(latentStyleExceptionInfo89);
            latentStyles1.Append(latentStyleExceptionInfo90);
            latentStyles1.Append(latentStyleExceptionInfo91);
            latentStyles1.Append(latentStyleExceptionInfo92);
            latentStyles1.Append(latentStyleExceptionInfo93);
            latentStyles1.Append(latentStyleExceptionInfo94);
            latentStyles1.Append(latentStyleExceptionInfo95);
            latentStyles1.Append(latentStyleExceptionInfo96);
            latentStyles1.Append(latentStyleExceptionInfo97);
            latentStyles1.Append(latentStyleExceptionInfo98);
            latentStyles1.Append(latentStyleExceptionInfo99);
            latentStyles1.Append(latentStyleExceptionInfo100);
            latentStyles1.Append(latentStyleExceptionInfo101);
            latentStyles1.Append(latentStyleExceptionInfo102);
            latentStyles1.Append(latentStyleExceptionInfo103);
            latentStyles1.Append(latentStyleExceptionInfo104);
            latentStyles1.Append(latentStyleExceptionInfo105);
            latentStyles1.Append(latentStyleExceptionInfo106);
            latentStyles1.Append(latentStyleExceptionInfo107);
            latentStyles1.Append(latentStyleExceptionInfo108);
            latentStyles1.Append(latentStyleExceptionInfo109);
            latentStyles1.Append(latentStyleExceptionInfo110);
            latentStyles1.Append(latentStyleExceptionInfo111);
            latentStyles1.Append(latentStyleExceptionInfo112);
            latentStyles1.Append(latentStyleExceptionInfo113);
            latentStyles1.Append(latentStyleExceptionInfo114);
            latentStyles1.Append(latentStyleExceptionInfo115);
            latentStyles1.Append(latentStyleExceptionInfo116);
            latentStyles1.Append(latentStyleExceptionInfo117);
            latentStyles1.Append(latentStyleExceptionInfo118);
            latentStyles1.Append(latentStyleExceptionInfo119);
            latentStyles1.Append(latentStyleExceptionInfo120);
            latentStyles1.Append(latentStyleExceptionInfo121);
            latentStyles1.Append(latentStyleExceptionInfo122);
            latentStyles1.Append(latentStyleExceptionInfo123);
            latentStyles1.Append(latentStyleExceptionInfo124);
            latentStyles1.Append(latentStyleExceptionInfo125);
            latentStyles1.Append(latentStyleExceptionInfo126);
            latentStyles1.Append(latentStyleExceptionInfo127);
            latentStyles1.Append(latentStyleExceptionInfo128);
            latentStyles1.Append(latentStyleExceptionInfo129);
            latentStyles1.Append(latentStyleExceptionInfo130);
            latentStyles1.Append(latentStyleExceptionInfo131);
            latentStyles1.Append(latentStyleExceptionInfo132);
            latentStyles1.Append(latentStyleExceptionInfo133);
            latentStyles1.Append(latentStyleExceptionInfo134);
            latentStyles1.Append(latentStyleExceptionInfo135);
            latentStyles1.Append(latentStyleExceptionInfo136);
            latentStyles1.Append(latentStyleExceptionInfo137);

            Style style1 = new Style() { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true };
            StyleName styleName1 = new StyleName() { Val = "Normal" };
            PrimaryStyle primaryStyle1 = new PrimaryStyle();
            Rsid rsid12 = new Rsid() { Val = "003D3FA0" };

            style1.Append(styleName1);
            style1.Append(primaryStyle1);
            style1.Append(rsid12);

            Style style2 = new Style() { Type = StyleValues.Character, StyleId = "Fuentedeprrafopredeter", Default = true };
            StyleName styleName2 = new StyleName() { Val = "Default Paragraph Font" };
            UIPriority uIPriority1 = new UIPriority() { Val = 1 };
            SemiHidden semiHidden1 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed1 = new UnhideWhenUsed();

            style2.Append(styleName2);
            style2.Append(uIPriority1);
            style2.Append(semiHidden1);
            style2.Append(unhideWhenUsed1);

            Style style3 = new Style() { Type = StyleValues.Table, StyleId = "Tablanormal", Default = true };
            StyleName styleName3 = new StyleName() { Val = "Normal Table" };
            UIPriority uIPriority2 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden2 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed2 = new UnhideWhenUsed();
            PrimaryStyle primaryStyle2 = new PrimaryStyle();

            StyleTableProperties styleTableProperties1 = new StyleTableProperties();
            TableIndentation tableIndentation1 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableCellMarginDefault tableCellMarginDefault1 = new TableCellMarginDefault();
            TopMargin topMargin1 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin1 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin1 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin1 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault1.Append(topMargin1);
            tableCellMarginDefault1.Append(tableCellLeftMargin1);
            tableCellMarginDefault1.Append(bottomMargin1);
            tableCellMarginDefault1.Append(tableCellRightMargin1);

            styleTableProperties1.Append(tableIndentation1);
            styleTableProperties1.Append(tableCellMarginDefault1);

            style3.Append(styleName3);
            style3.Append(uIPriority2);
            style3.Append(semiHidden2);
            style3.Append(unhideWhenUsed2);
            style3.Append(primaryStyle2);
            style3.Append(styleTableProperties1);

            Style style4 = new Style() { Type = StyleValues.Numbering, StyleId = "Sinlista", Default = true };
            StyleName styleName4 = new StyleName() { Val = "No List" };
            UIPriority uIPriority3 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden3 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed3 = new UnhideWhenUsed();

            style4.Append(styleName4);
            style4.Append(uIPriority3);
            style4.Append(semiHidden3);
            style4.Append(unhideWhenUsed3);

            Style style5 = new Style() { Type = StyleValues.Table, StyleId = "Tablaconcuadrcula" };
            StyleName styleName5 = new StyleName() { Val = "Table Grid" };
            BasedOn basedOn1 = new BasedOn() { Val = "Tablanormal" };
            UIPriority uIPriority4 = new UIPriority() { Val = 39 };
            Rsid rsid13 = new Rsid() { Val = "00E553F0" };

            StyleParagraphProperties styleParagraphProperties1 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines2 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties1.Append(spacingBetweenLines2);

            StyleTableProperties styleTableProperties2 = new StyleTableProperties();
            TableIndentation tableIndentation2 = new TableIndentation() { Width = 0, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.Single, Color = "auto", Size = (UInt32Value)4U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);

            TableCellMarginDefault tableCellMarginDefault2 = new TableCellMarginDefault();
            TopMargin topMargin2 = new TopMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellLeftMargin tableCellLeftMargin2 = new TableCellLeftMargin() { Width = 108, Type = TableWidthValues.Dxa };
            BottomMargin bottomMargin2 = new BottomMargin() { Width = "0", Type = TableWidthUnitValues.Dxa };
            TableCellRightMargin tableCellRightMargin2 = new TableCellRightMargin() { Width = 108, Type = TableWidthValues.Dxa };

            tableCellMarginDefault2.Append(topMargin2);
            tableCellMarginDefault2.Append(tableCellLeftMargin2);
            tableCellMarginDefault2.Append(bottomMargin2);
            tableCellMarginDefault2.Append(tableCellRightMargin2);

            styleTableProperties2.Append(tableIndentation2);
            styleTableProperties2.Append(tableBorders1);
            styleTableProperties2.Append(tableCellMarginDefault2);

            style5.Append(styleName5);
            style5.Append(basedOn1);
            style5.Append(uIPriority4);
            style5.Append(rsid13);
            style5.Append(styleParagraphProperties1);
            style5.Append(styleTableProperties2);

            Style style6 = new Style() { Type = StyleValues.Paragraph, StyleId = "Textodeglobo" };
            StyleName styleName6 = new StyleName() { Val = "Balloon Text" };
            BasedOn basedOn2 = new BasedOn() { Val = "Normal" };
            LinkedStyle linkedStyle1 = new LinkedStyle() { Val = "TextodegloboCar" };
            UIPriority uIPriority5 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden4 = new SemiHidden();
            UnhideWhenUsed unhideWhenUsed4 = new UnhideWhenUsed();
            Rsid rsid14 = new Rsid() { Val = "00E553F0" };

            StyleParagraphProperties styleParagraphProperties2 = new StyleParagraphProperties();
            SpacingBetweenLines spacingBetweenLines3 = new SpacingBetweenLines() { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto };

            styleParagraphProperties2.Append(spacingBetweenLines3);

            StyleRunProperties styleRunProperties1 = new StyleRunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize3 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties1.Append(runFonts2);
            styleRunProperties1.Append(fontSize3);
            styleRunProperties1.Append(fontSizeComplexScript2);

            style6.Append(styleName6);
            style6.Append(basedOn2);
            style6.Append(linkedStyle1);
            style6.Append(uIPriority5);
            style6.Append(semiHidden4);
            style6.Append(unhideWhenUsed4);
            style6.Append(rsid14);
            style6.Append(styleParagraphProperties2);
            style6.Append(styleRunProperties1);

            Style style7 = new Style() { Type = StyleValues.Character, StyleId = "TextodegloboCar", CustomStyle = true };
            StyleName styleName7 = new StyleName() { Val = "Texto de globo Car" };
            BasedOn basedOn3 = new BasedOn() { Val = "Fuentedeprrafopredeter" };
            LinkedStyle linkedStyle2 = new LinkedStyle() { Val = "Textodeglobo" };
            UIPriority uIPriority6 = new UIPriority() { Val = 99 };
            SemiHidden semiHidden5 = new SemiHidden();
            Rsid rsid15 = new Rsid() { Val = "00E553F0" };

            StyleRunProperties styleRunProperties2 = new StyleRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Tahoma", HighAnsi = "Tahoma", ComplexScript = "Tahoma" };
            FontSize fontSize4 = new FontSize() { Val = "16" };
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "16" };

            styleRunProperties2.Append(runFonts3);
            styleRunProperties2.Append(fontSize4);
            styleRunProperties2.Append(fontSizeComplexScript3);

            style7.Append(styleName7);
            style7.Append(basedOn3);
            style7.Append(linkedStyle2);
            style7.Append(uIPriority6);
            style7.Append(semiHidden5);
            style7.Append(rsid15);
            style7.Append(styleRunProperties2);

            styles1.Append(docDefaults1);
            styles1.Append(latentStyles1);
            styles1.Append(style1);
            styles1.Append(style2);
            styles1.Append(style3);
            styles1.Append(style4);
            styles1.Append(style5);
            styles1.Append(style6);
            styles1.Append(style7);

            styleDefinitionsPart1.Styles = styles1;
        }

        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Tema de Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex2);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex3);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex4);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex5);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex6);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex7);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex8);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex9);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex10);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor1.Append(rgbColorModelHex11);

            colorScheme1.Append(dark1Color1);
            colorScheme1.Append(light1Color1);
            colorScheme1.Append(dark2Color1);
            colorScheme1.Append(light2Color1);
            colorScheme1.Append(accent1Color1);
            colorScheme1.Append(accent2Color1);
            colorScheme1.Append(accent3Color1);
            colorScheme1.Append(accent4Color1);
            colorScheme1.Append(accent5Color1);
            colorScheme1.Append(accent6Color1);
            colorScheme1.Append(hyperlink1);
            colorScheme1.Append(followedHyperlinkColor1);

            A.FontScheme fontScheme1 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Calibri Light" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ ゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Angsana New" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            majorFont1.Append(latinFont1);
            majorFont1.Append(eastAsianFont1);
            majorFont1.Append(complexScriptFont1);
            majorFont1.Append(supplementalFont1);
            majorFont1.Append(supplementalFont2);
            majorFont1.Append(supplementalFont3);
            majorFont1.Append(supplementalFont4);
            majorFont1.Append(supplementalFont5);
            majorFont1.Append(supplementalFont6);
            majorFont1.Append(supplementalFont7);
            majorFont1.Append(supplementalFont8);
            majorFont1.Append(supplementalFont9);
            majorFont1.Append(supplementalFont10);
            majorFont1.Append(supplementalFont11);
            majorFont1.Append(supplementalFont12);
            majorFont1.Append(supplementalFont13);
            majorFont1.Append(supplementalFont14);
            majorFont1.Append(supplementalFont15);
            majorFont1.Append(supplementalFont16);
            majorFont1.Append(supplementalFont17);
            majorFont1.Append(supplementalFont18);
            majorFont1.Append(supplementalFont19);
            majorFont1.Append(supplementalFont20);
            majorFont1.Append(supplementalFont21);
            majorFont1.Append(supplementalFont22);
            majorFont1.Append(supplementalFont23);
            majorFont1.Append(supplementalFont24);
            majorFont1.Append(supplementalFont25);
            majorFont1.Append(supplementalFont26);
            majorFont1.Append(supplementalFont27);
            majorFont1.Append(supplementalFont28);
            majorFont1.Append(supplementalFont29);
            majorFont1.Append(supplementalFont30);

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ 明朝" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Cordia New" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont31);
            minorFont1.Append(supplementalFont32);
            minorFont1.Append(supplementalFont33);
            minorFont1.Append(supplementalFont34);
            minorFont1.Append(supplementalFont35);
            minorFont1.Append(supplementalFont36);
            minorFont1.Append(supplementalFont37);
            minorFont1.Append(supplementalFont38);
            minorFont1.Append(supplementalFont39);
            minorFont1.Append(supplementalFont40);
            minorFont1.Append(supplementalFont41);
            minorFont1.Append(supplementalFont42);
            minorFont1.Append(supplementalFont43);
            minorFont1.Append(supplementalFont44);
            minorFont1.Append(supplementalFont45);
            minorFont1.Append(supplementalFont46);
            minorFont1.Append(supplementalFont47);
            minorFont1.Append(supplementalFont48);
            minorFont1.Append(supplementalFont49);
            minorFont1.Append(supplementalFont50);
            minorFont1.Append(supplementalFont51);
            minorFont1.Append(supplementalFont52);
            minorFont1.Append(supplementalFont53);
            minorFont1.Append(supplementalFont54);
            minorFont1.Append(supplementalFont55);
            minorFont1.Append(supplementalFont56);
            minorFont1.Append(supplementalFont57);
            minorFont1.Append(supplementalFont58);
            minorFont1.Append(supplementalFont59);
            minorFont1.Append(supplementalFont60);

            fontScheme1.Append(majorFont1);
            fontScheme1.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill2);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill3);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill4);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill5);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex12.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex12);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill6 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill6.Append(schemeColor11);

            A.SolidFill solidFill7 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill7.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill6);
            backgroundFillStyleList1.Append(solidFill7);
            backgroundFillStyleList1.Append(gradientFill3);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme1);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList1 = new A.OfficeStyleSheetExtensionList();

            A.OfficeStyleSheetExtension officeStyleSheetExtension1 = new A.OfficeStyleSheetExtension() { Uri = "{05A4C25C-085E-4340-85A3-A5531E510DB2}" };

            Thm15.ThemeFamily themeFamily1 = new Thm15.ThemeFamily() { Name = "Office Theme", Id = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}", Vid = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" };
            themeFamily1.AddNamespaceDeclaration("thm15", "http://schemas.microsoft.com/office/thememl/2012/main");

            officeStyleSheetExtension1.Append(themeFamily1);

            officeStyleSheetExtensionList1.Append(officeStyleSheetExtension1);

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);
            theme1.Append(officeStyleSheetExtensionList1);

            themePart1.Theme = theme1;
        }

        private void GenerateFontTablePart1Content(FontTablePart fontTablePart1)
        {
            Fonts fonts1 = new Fonts();
            fonts1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            fonts1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            Font font1 = new Font() { Name = "Calibri" };
            Panose1Number panose1Number1 = new Panose1Number() { Val = "020F0502020204030204" };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily1 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch1 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature1 = new FontSignature() { UnicodeSignature0 = "E00002FF", UnicodeSignature1 = "4000ACFF", UnicodeSignature2 = "00000001", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font1.Append(panose1Number1);
            font1.Append(fontCharSet1);
            font1.Append(fontFamily1);
            font1.Append(pitch1);
            font1.Append(fontSignature1);

            Font font2 = new Font() { Name = "Times New Roman" };
            Panose1Number panose1Number2 = new Panose1Number() { Val = "02020603050405020304" };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily2 = new FontFamily() { Val = FontFamilyValues.Roman };
            Pitch pitch2 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature2 = new FontSignature() { UnicodeSignature0 = "E0002EFF", UnicodeSignature1 = "C0007843", UnicodeSignature2 = "00000009", UnicodeSignature3 = "00000000", CodePageSignature0 = "000001FF", CodePageSignature1 = "00000000" };

            font2.Append(panose1Number2);
            font2.Append(fontCharSet2);
            font2.Append(fontFamily2);
            font2.Append(pitch2);
            font2.Append(fontSignature2);

            Font font3 = new Font() { Name = "Tahoma" };
            Panose1Number panose1Number3 = new Panose1Number() { Val = "020B0604030504040204" };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily3 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch3 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature3 = new FontSignature() { UnicodeSignature0 = "E1002EFF", UnicodeSignature1 = "C000605B", UnicodeSignature2 = "00000029", UnicodeSignature3 = "00000000", CodePageSignature0 = "000101FF", CodePageSignature1 = "00000000" };

            font3.Append(panose1Number3);
            font3.Append(fontCharSet3);
            font3.Append(fontFamily3);
            font3.Append(pitch3);
            font3.Append(fontSignature3);

            Font font4 = new Font() { Name = "Calibri Light" };
            Panose1Number panose1Number4 = new Panose1Number() { Val = "020F0302020204030204" };
            FontCharSet fontCharSet4 = new FontCharSet() { Val = "00" };
            FontFamily fontFamily4 = new FontFamily() { Val = FontFamilyValues.Swiss };
            Pitch pitch4 = new Pitch() { Val = FontPitchValues.Variable };
            FontSignature fontSignature4 = new FontSignature() { UnicodeSignature0 = "A00002EF", UnicodeSignature1 = "4000207B", UnicodeSignature2 = "00000000", UnicodeSignature3 = "00000000", CodePageSignature0 = "0000019F", CodePageSignature1 = "00000000" };

            font4.Append(panose1Number4);
            font4.Append(fontCharSet4);
            font4.Append(fontFamily4);
            font4.Append(pitch4);
            font4.Append(fontSignature4);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            fontTablePart1.Fonts = fonts1;
        }

        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Lduenas";
            document.PackageProperties.Revision = "2";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2022-04-26T23:47:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2022-04-26T23:53:00Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Lduenas";
        }

        private string imagePart1Data = "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";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

    }
}
