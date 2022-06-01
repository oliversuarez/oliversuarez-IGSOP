using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using System;
using System.IO;
using System.Data;
using System.Collections.Generic;

namespace General.Librerias.CodigoUsuario
{
    public class ExcelSDK
    {
        string titulo;
        string usuario;
        string sede;
        DataTable tabla;
        Dictionary<string, string> piePagina = null;
        int nCols;
        int nFilas;
        public void CrearArchivo(string archivoExcel, string _titulo, string _usuario, string _sede, DataTable _tabla, Dictionary<string, string> _piePagina = null)
        {
            titulo = _titulo;
            usuario = _usuario;
            sede = _sede;
            tabla = _tabla;
            piePagina = _piePagina;
            nCols = tabla.Columns.Count;
            nFilas = tabla.Rows.Count;            
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(archivoExcel, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        public byte[] CrearBytes(string _titulo, string _usuario, string _sede, DataTable _tabla, Dictionary<string, string> _piePagina = null)
        {
            byte[] rpta = null;
            titulo = _titulo;
            usuario = _usuario;
            sede = _sede;
            tabla = _tabla;
            piePagina = _piePagina;
            nCols = tabla.Columns.Count;
            nFilas = tabla.Rows.Count;
            using (MemoryStream ms = new MemoryStream())
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook))
                {
                    CreateParts(package);
                }
                rpta = ms.ToArray();
            }
            return rpta;
        }

        private void CreateParts(SpreadsheetDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId3");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId2");
            GenerateThemePart1Content(themePart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            DrawingsPart drawingsPart1 = worksheetPart1.AddNewPart<DrawingsPart>("rId2");
            GenerateDrawingsPart1Content(drawingsPart1);

            ImagePart imagePart1 = drawingsPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
            GenerateImagePart1Content(imagePart1);

            SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1 = worksheetPart1.AddNewPart<SpreadsheetPrinterSettingsPart>("rId1");
            GenerateSpreadsheetPrinterSettingsPart1Content(spreadsheetPrinterSettingsPart1);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId4");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            SetPackageProperties(document);
        }

        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity1 = new Ap.DocumentSecurity();
            documentSecurity1.Text = "0";
            Ap.ScaleCrop scaleCrop1 = new Ap.ScaleCrop();
            scaleCrop1.Text = "false";

            Ap.HeadingPairs headingPairs1 = new Ap.HeadingPairs();

            Vt.VTVector vTVector1 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant1 = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR1 = new Vt.VTLPSTR();
            vTLPSTR1.Text = "Hojas de cálculo";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Reporte";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "12.0000";

            properties1.Append(application1);
            properties1.Append(documentSecurity1);
            properties1.Append(scaleCrop1);
            properties1.Append(headingPairs1);
            properties1.Append(titlesOfParts1);
            properties1.Append(linksUpToDate1);
            properties1.Append(sharedDocument1);
            properties1.Append(hyperlinksChanged1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "4", LowestEdited = "4", BuildVersion = "4507" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 240, YWindow = 30, WindowWidth = (UInt32Value)20115U, WindowHeight = (UInt32Value)8010U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet() { Name = "Reporte", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)125725U };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);

            workbookPart1.Workbook = workbook1;
        }

        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet();

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)4U };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 14D };
            Color color3 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName3 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };

            font3.Append(bold2);
            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);

            Font font4 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = 11D };
            Color color4 = new Color() { Rgb = "FF00B050" };
            FontName fontName4 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font4.Append(bold3);
            font4.Append(fontSize4);
            font4.Append(color4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering4);
            font4.Append(fontScheme3);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)2U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color5);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color6);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color7 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color7);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color8);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            borders1.Append(border1);
            borders1.Append(border2);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)7U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat3.Append(alignment1);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left };

            cellFormat4.Append(alignment2);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Right };

            cellFormat5.Append(alignment3);
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)14U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyBorder = true };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);
            cellFormats1.Append(cellFormat8);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium9", DefaultPivotStyle = "PivotStyleLight16" };

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);

            workbookStylesPart1.Stylesheet = stylesheet1;
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
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "1F497D" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "EEECE1" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "4F81BD" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "C0504D" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "9BBB59" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "8064A2" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4BACC6" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "F79646" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0000FF" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "800080" };

            followedHyperlinkColor1.Append(rgbColorModelHex10);

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

            A.FontScheme fontScheme4 = new A.FontScheme() { Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont() { Typeface = "Cambria" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
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

            A.MinorFont minorFont1 = new A.MinorFont();
            A.LatinFont latinFont2 = new A.LatinFont() { Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont() { Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont() { Typeface = "" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" };

            minorFont1.Append(latinFont2);
            minorFont1.Append(eastAsianFont2);
            minorFont1.Append(complexScriptFont2);
            minorFont1.Append(supplementalFont30);
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

            fontScheme4.Append(majorFont1);
            fontScheme4.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint1 = new A.Tint() { Val = 50000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 300000 };

            schemeColor2.Append(tint1);
            schemeColor2.Append(saturationModulation1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 35000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint2 = new A.Tint() { Val = 37000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 300000 };

            schemeColor3.Append(tint2);
            schemeColor3.Append(saturationModulation2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint3 = new A.Tint() { Val = 15000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 350000 };

            schemeColor4.Append(tint3);
            schemeColor4.Append(saturationModulation3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 16200000, Scaled = true };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade1 = new A.Shade() { Val = 51000 };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 130000 };

            schemeColor5.Append(shade1);
            schemeColor5.Append(saturationModulation4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 80000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade2 = new A.Shade() { Val = 93000 };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 130000 };

            schemeColor6.Append(shade2);
            schemeColor6.Append(saturationModulation5);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade3 = new A.Shade() { Val = 94000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 135000 };

            schemeColor7.Append(shade3);
            schemeColor7.Append(saturationModulation6);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 16200000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 9525, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();

            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade4 = new A.Shade() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 105000 };

            schemeColor8.Append(shade4);
            schemeColor8.Append(saturationModulation7);

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);

            A.Outline outline2 = new A.Outline() { Width = 25400, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);

            A.Outline outline3 = new A.Outline() { Width = 38100, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);

            lineStyleList1.Append(outline1);
            lineStyleList1.Append(outline2);
            lineStyleList1.Append(outline3);

            A.EffectStyleList effectStyleList1 = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();

            A.EffectList effectList1 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 38000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList1.Append(outerShadow1);

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();

            A.EffectList effectList2 = new A.EffectList();

            A.OuterShadow outerShadow2 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex12 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha2 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex12.Append(alpha2);

            outerShadow2.Append(rgbColorModelHex12);

            effectList2.Append(outerShadow2);

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow3 = new A.OuterShadow() { BlurRadius = 40000L, Distance = 23000L, Direction = 5400000, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex13 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha3 = new A.Alpha() { Val = 35000 };

            rgbColorModelHex13.Append(alpha3);

            outerShadow3.Append(rgbColorModelHex13);

            effectList3.Append(outerShadow3);

            A.Scene3DType scene3DType1 = new A.Scene3DType();

            A.Camera camera1 = new A.Camera() { Preset = A.PresetCameraValues.OrthographicFront };
            A.Rotation rotation1 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 0 };

            camera1.Append(rotation1);

            A.LightRig lightRig1 = new A.LightRig() { Rig = A.LightRigValues.ThreePoints, Direction = A.LightRigDirectionValues.Top };
            A.Rotation rotation2 = new A.Rotation() { Latitude = 0, Longitude = 0, Revolution = 1200000 };

            lightRig1.Append(rotation2);

            scene3DType1.Append(camera1);
            scene3DType1.Append(lightRig1);

            A.Shape3DType shape3DType1 = new A.Shape3DType();
            A.BevelTop bevelTop1 = new A.BevelTop() { Width = 63500L, Height = 25400L };

            shape3DType1.Append(bevelTop1);

            effectStyle3.Append(effectList3);
            effectStyle3.Append(scene3DType1);
            effectStyle3.Append(shape3DType1);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint4 = new A.Tint() { Val = 40000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 350000 };

            schemeColor12.Append(tint4);
            schemeColor12.Append(saturationModulation8);

            gradientStop7.Append(schemeColor12);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 40000 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 45000 };
            A.Shade shade5 = new A.Shade() { Val = 99000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 350000 };

            schemeColor13.Append(tint5);
            schemeColor13.Append(shade5);
            schemeColor13.Append(saturationModulation9);

            gradientStop8.Append(schemeColor13);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade6 = new A.Shade() { Val = 20000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 255000 };

            schemeColor14.Append(shade6);
            schemeColor14.Append(saturationModulation10);

            gradientStop9.Append(schemeColor14);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);

            A.PathGradientFill pathGradientFill1 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle1 = new A.FillToRectangle() { Left = 50000, Top = -80000, Right = 50000, Bottom = 180000 };

            pathGradientFill1.Append(fillToRectangle1);

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(pathGradientFill1);

            A.GradientFill gradientFill4 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList4 = new A.GradientStopList();

            A.GradientStop gradientStop10 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 80000 };
            A.SaturationModulation saturationModulation11 = new A.SaturationModulation() { Val = 300000 };

            schemeColor15.Append(tint6);
            schemeColor15.Append(saturationModulation11);

            gradientStop10.Append(schemeColor15);

            A.GradientStop gradientStop11 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor16 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade7 = new A.Shade() { Val = 30000 };
            A.SaturationModulation saturationModulation12 = new A.SaturationModulation() { Val = 200000 };

            schemeColor16.Append(shade7);
            schemeColor16.Append(saturationModulation12);

            gradientStop11.Append(schemeColor16);

            gradientStopList4.Append(gradientStop10);
            gradientStopList4.Append(gradientStop11);

            A.PathGradientFill pathGradientFill2 = new A.PathGradientFill() { Path = A.PathShadeValues.Circle };
            A.FillToRectangle fillToRectangle2 = new A.FillToRectangle() { Left = 50000, Top = 50000, Right = 50000, Bottom = 50000 };

            pathGradientFill2.Append(fillToRectangle2);

            gradientFill4.Append(gradientStopList4);
            gradientFill4.Append(pathGradientFill2);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(gradientFill3);
            backgroundFillStyleList1.Append(gradientFill4);

            formatScheme1.Append(fillStyleList1);
            formatScheme1.Append(lineStyleList1);
            formatScheme1.Append(effectStyleList1);
            formatScheme1.Append(backgroundFillStyleList1);

            themeElements1.Append(colorScheme1);
            themeElements1.Append(fontScheme4);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet();
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A2:G6" };

            SheetViews sheetViews1 = new SheetViews();
            SheetView sheetView1 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { BaseColumnWidth = (UInt32Value)10U, DefaultRowHeight = 15D };

            double ancho;
            Columns columns1 = new Columns();
            for(int j = 0; j < nCols; j++)
            {
                ancho = double.Parse(tabla.Columns[j].Caption);
                Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = DoubleValue.FromDouble(ancho), CustomWidth = true };
                columns1.Append(column1);
            }
            SheetData sheetData1 = new SheetData();
            string rangoCols = string.Format("1:{0}",nCols);
            int nCol;

            //Crear Fila 0 del Logo
            Row row0 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = rangoCols }, Height = 52.5D, CustomHeight = true };

            //Crear las celdas vacias de la fila 0
            for (int j = 0; j < nCols; j++)
            {
                nCol = 65 + j;
                Cell cell01 = new Cell() { CellReference = string.Format("{0}1", (char)(nCol)), StyleIndex = (UInt32Value)1U };
                row0.Append(cell01);
            }
            sheetData1.Append(row0);

            //Crear Fila 1 del Titulo
            Row row1 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = rangoCols }, Height = 18D };
            //Crear la celda con el Titulo
            Cell cell11 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "0";
            cell11.Append(cellValue11);
            row1.Append(cell11);

            //Crear las celdas vacias de la fila 1
            for (int j = 1; j < nCols; j++)
            {
                nCol = 65 + j;
                Cell cell12 = new Cell() { CellReference = string.Format("{0}2", (char)(nCol)), StyleIndex = (UInt32Value)1U };
                row1.Append(cell12);
            }
            sheetData1.Append(row1);
            
            //Crear Fila 2 de 2 Subtitulos: izquierda y derecha
            Row row2 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:7" } };
            //Dividir la cantidad de columnas en 2 para tener 2 agrupaciones
            int nMitad = (int)(nCols / 2);
            //Crear Celda con Subtitulo 1 a la izquierda
            Cell cell21 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "1";
            cell21.Append(cellValue21);
            row2.Append(cell21);
            //Crear las celdas vacias del subtitulo 1 de la fila 2
            for (int j = 1; j < nMitad; j++)
            {
                nCol = 65 + j;
                Cell cell22 = new Cell() { CellReference = string.Format("{0}3", (char)(nCol)), StyleIndex = (UInt32Value)2U };
                row2.Append(cell22);
            }
            //Crear Celda con Subtitulo 2 a la derecha
            Cell cell23 = new Cell() { CellReference = string.Format("{0}3",(char)(nMitad+65)), StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "2";
            cell23.Append(cellValue23);
            row2.Append(cell23);            
            //Crear las celdas vacias del subtitulo 2 de la fila 2
            for (int j = nMitad; j < nCols; j++)
            {
                nCol = 66 + j;
                Cell cell24 = new Cell() { CellReference = string.Format("{0}3", (char)(nCol)), StyleIndex = (UInt32Value)2U };
                row2.Append(cell24);
            }
            sheetData1.Append(row2);
            //Crear Fila 3 con los nombres de campos
            Row row3 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:7" } };
            for (int j = 0; j < nCols; j++)
            {
                nCol = 65 + j;
                Cell cell31 = new Cell() { CellReference = string.Format("{0}4", (char)(nCol)), StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
                CellValue cellValue31 = new CellValue();
                cellValue31.Text = (3+j).ToString();
                cell31.Append(cellValue31);
                row3.Append(cell31);
            }
            sheetData1.Append(row3);
            
            //Crear las Filas con la data de la tabla
            string tipoDato;
            UInt32Value estiloCelda = (UInt32Value)4U;
            CellValues tipoCelda = CellValues.SharedString;
            string valorCelda = "";
            int cString = 3 + nCols;
            for (int i = 0; i < nFilas; i++)
            {
                Row row4 = new Row() { RowIndex = Convert.ToUInt32(5+i), Spans = new ListValue<StringValue>() { InnerText = rangoCols } };
                for (int j = 0; j < nCols; j++)
                {
                    tipoDato = tabla.Columns[j].DataType.ToString();
                    if (tipoDato.Contains("String"))
                    {
                        estiloCelda = (UInt32Value)4U;
                        tipoCelda = CellValues.SharedString;
                        valorCelda = cString.ToString();
                        cString++;
                    }
                    else if (tipoDato.Contains("Int") || tipoDato.Contains("Decimal"))
                    {
                        estiloCelda = (UInt32Value)4U;
                        tipoCelda = CellValues.Number;
                        valorCelda = tabla.Rows[i][j].ToString();
                    }
                    nCol = 65 + j;
                    Cell cell41 = new Cell() { CellReference = string.Format("{0}{1}", (char)(nCol),5+i), StyleIndex = estiloCelda, DataType = tipoCelda };
                    CellValue cellValue41 = new CellValue();
                    cellValue41.Text = valorCelda;
                    cell41.Append(cellValue41);
                    row4.Append(cell41);
                }
                sheetData1.Append(row4);
            }
            
            //Crear la Fila con los Pies de Paginas
            if (piePagina != null && piePagina.Count > 0)
            {
                Row row5 = new Row() { RowIndex = Convert.ToUInt32(5 + nFilas), Spans = new ListValue<StringValue>() { InnerText = rangoCols } };
                foreach (KeyValuePair<string, string> item in piePagina)
                {
                    Cell cell51 = new Cell() { CellReference = string.Format("{0}{1}", item.Key, 5 + nFilas), StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
                    CellValue cellValue51 = new CellValue();
                    cellValue51.Text = cString.ToString();
                    cString++;
                    cell51.Append(cellValue51);
                    row5.Append(cell51);
                }
                sheetData1.Append(row5);
            }

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)3U };
            MergeCell mergeCell0 = new MergeCell() { Reference = string.Format("A1:{0}1", (char)(64 + nCols)) };
            MergeCell mergeCell1 = new MergeCell() { Reference = string.Format("A2:{0}2", (char)(64 + nCols)) };
            MergeCell mergeCell2 = new MergeCell() { Reference = string.Format("A3:{0}3", (char)(64 + nMitad)) };
            MergeCell mergeCell3 = new MergeCell() { Reference = string.Format("{0}3:{1}3", (char)(65 + nMitad), (char)(64 + nCols)) };

            mergeCells1.Append(mergeCell0);
            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);

            PageMargins pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, Orientation = OrientationValues.Portrait, HorizontalDpi = (UInt32Value)1U, VerticalDpi = (UInt32Value)1U, Id = "rId1" };
            
            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            //worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(pageMargins1);
            worksheet1.Append(pageSetup1);
            
            //Para el Grafico
            Drawing drawing1 = new Drawing() { Id = "rId2" };
            worksheet1.Append(drawing1);
            
            worksheetPart1.Worksheet = worksheet1;
        }

        private void GenerateDrawingsPart1Content(DrawingsPart drawingsPart1)
        {
            Xdr.WorksheetDrawing worksheetDrawing1 = new Xdr.WorksheetDrawing();
            worksheetDrawing1.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            worksheetDrawing1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Xdr.TwoCellAnchor twoCellAnchor1 = new Xdr.TwoCellAnchor();

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = "0";
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = "0";
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            columnId2.Text = "2";
            Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            columnOffset2.Text = "167640";
            Xdr.RowId rowId2 = new Xdr.RowId();
            rowId2.Text = "0";
            Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            rowOffset2.Text = "647700";

            toMarker1.Append(columnId2);
            toMarker1.Append(columnOffset2);
            toMarker1.Append(rowId2);
            toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();

            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = (UInt32Value)3U, Name = "Picture 1" };

            A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

            A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{CB2255AE-3FF9-453F-B440-93DF0C0596EF}\" />");

            nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

            nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

            nonVisualDrawingProperties1.Append(nonVisualDrawingPropertiesExtensionList1);

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId1", CompressionState = A.BlipCompressionValues.Print };
            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

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

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties() { BlackWhiteMode = A.BlackWhiteModeValues.Auto };

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
            hiddenFillProperties1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            A.SolidFill solidFill6 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex14 = new A.RgbColorModelHex() { Val = "FFFFFF" };

            solidFill6.Append(rgbColorModelHex14);

            hiddenFillProperties1.Append(solidFill6);

            shapePropertiesExtension1.Append(hiddenFillProperties1);

            shapePropertiesExtensionList1.Append(shapePropertiesExtension1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(shapePropertiesExtensionList1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            twoCellAnchor1.Append(fromMarker1);
            twoCellAnchor1.Append(toMarker1);
            twoCellAnchor1.Append(picture1);
            twoCellAnchor1.Append(clientData1);

            worksheetDrawing1.Append(twoCellAnchor1);

            drawingsPart1.WorksheetDrawing = worksheetDrawing1;
        }

        private void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        private void GenerateSpreadsheetPrinterSettingsPart1Content(SpreadsheetPrinterSettingsPart spreadsheetPrinterSettingsPart1)
        {
            System.IO.Stream data = GetBinaryDataStream(spreadsheetPrinterSettingsPart1Data);
            spreadsheetPrinterSettingsPart1.FeedData(data);
            data.Close();
        }

        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)16U, UniqueCount = (UInt32Value)16U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = titulo;
            sharedStringItem1.Append(text1);
            sharedStringTable1.Append(sharedStringItem1);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = usuario;
            sharedStringItem2.Append(text2);
            sharedStringTable1.Append(sharedStringItem2);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = sede;
            sharedStringItem3.Append(text3);
            sharedStringTable1.Append(sharedStringItem3);

            for (int j = 0; j < nCols; j++)
            {
                SharedStringItem sharedStringItem4 = new SharedStringItem();
                Text text4 = new Text();
                text4.Text = tabla.Columns[j].ColumnName;
                sharedStringItem4.Append(text4);
                sharedStringTable1.Append(sharedStringItem4);
            }
            string tipoDato = "";
            string valor;
            for (int i = 0; i < nFilas; i++)
            {
                for (int j = 0; j < nCols; j++)
                {
                    tipoDato = tabla.Columns[j].DataType.ToString();
                    if (tipoDato.Contains("String"))
                    {
                        valor = (string)tabla.Rows[i][j];
                        SharedStringItem sharedStringItem5 = new SharedStringItem();
                        Text text5 = new Text();
                        text5.Text = valor;
                        sharedStringItem5.Append(text5);
                        sharedStringTable1.Append(sharedStringItem5);
                    }
                }
            }
            if(piePagina!=null && piePagina.Count > 0)
            {
                foreach(KeyValuePair<string, string> item in piePagina)
                {
                    SharedStringItem sharedStringItem6 = new SharedStringItem();
                    Text text6 = new Text();
                    text6.Text = item.Value;
                    sharedStringItem6.Append(text6);
                    sharedStringTable1.Append(sharedStringItem6);
                }
            }
            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Lduenas";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2022-04-25T00:02:27Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2022-04-25T00:11:56Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Lduenas";
        }

        private string spreadsheetPrinterSettingsPart1Data = "RQBQAFMATwBOACAATAAzADEAMQAwACAAUwBlAHIAaQBlAHMAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEEAALcAHgOD5uABwEACQCaCzQIZAABAAUBaAECAAEAaAEAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAgAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHgOAAABAQEBAAIAAAEAAAAAAAAAAAAAADgAAACADQAAuA0AAEAAAAD4DQAAgAAAAAAAAAAAAAAAAwAJBEUAUABTAE8ATgAgAEwAMwAxADEAMAAgAFMAZQByAGkAZQBzAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAAAAAAAABAOAAAAAAAAAAAAAAEAAAACAAMAAAAiAGgBaAEFAQAAAAAJADQImgseAB4AHgAeADQImgs7A5EEAQAAABAAFgAAAAAAAAAAAAAAAAAAAAAAAAACAAAABgAAAAAAAAAAAAIAAAAAAgAAAQAAAAIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADAAAABAAAAGQAZAA0CJoLHgAeAB4AHgAJAAAAAAAAAAAAAAD//wAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIABAACAAAAAAAAAAAAAQAyADIA1P4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAD/AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQABAAAAEAAAAAkEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIACAAgAcIAAACAAEAAQAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAABAAAAAAAAAAAAAAAAAAIAAgAGAAIAAAAAAAAAEA7/////AgAAAAAAAwDgAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACAAMAAgAAAAAAAAAAAAMAMgBkAHD+AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAABAAIASQAAABAAAAAJBAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACIiIiIAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgBAAQAACAAAQAAAGAQAABAAAAAAAABAAAAAAAAAAAAAgAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIAAAAAAAEAAAAAAAAAAAAAGgAAAAALAB4AHgAAAAAAAAAAAAAAAAABAAABAAAAAIEAAAAAAAAAAAAAAAAAAAABAQEB";

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

        private string imagePart1Data = "/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQECAQEBAQEBAgICAgICAgICAgICAgICAgICAgICAgICAgICAgL/2wBDAQEBAQEBAQICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgL/wAARCABQAOYDAREAAhEBAxEB/8QAHwAAAgIDAQADAQAAAAAAAAAAAAgHCQUGCgQBAgML/8QAThAAAAYBAgQCBAYOBQwDAAAAAQIDBAUGBwAICRESExQVFiEidwojJDE4thcyNjc5OkFCUXJ4srS4GCUmYbcZGiczNEZYWXGRlZeh1NX/xAAcAQEAAgMBAQEAAAAAAAAAAAAABgcEBQgJAwL/xABIEQABAwICAwsICAUBCQAAAAAAAQIDBAUGEQcSIRMiMTU2NzhBdHWzCDJxc3aytLUJFCMkQkNhsRUWJTM0UkRRU2OBg5Gk0f/aAAwDAQACEQMRAD8A7xch5AqeKqRZ8jXuVLBU2mxDqessydu6dpRUQxJ3Hj9ZBkRRUSJl9o/QQwgUBHlyDWHcK+ltdFJUTu1Iomq+R+SrqtThXJqKuz9EMSvrqW2UclRM7Uiiar5H5Kuq1OFVRM1yT9EIlhsuQ+5vAb/Iuz7MWPZ09ljJEuO8keCPdaQnYY5YUhaz8SycNFxTBUgoOkwVRcoFP3ikOJSpqamS4uxBY5ZLTVwbsqPZDM9iyxMmYu1k0aOY5Nu9cmaObnrIi7EXVJckxBZJJrTVQK9yPbBO5qzQtmYuSslYjmO2OTVembXN4erJUr2S74LzlS7X/aNvOqdboW5+nuJZqrFMGCjKk5Tqh0xE7muMpJZ0BzA2P1qIguuk9ZHK7QEUxVTTp/RzpWqcQ3utw3iOnhpL1CsrNxVuVPX06ouaRterkeqx5u2Zsnh37UTKSNlf4Gx7W3a4VFlvkUVPdYtdNRrdWCsh61ja5XbUau1us5HsVHtzTPKM79hG37E8yt8w4UKsfDtpfkZScAc6ijCuC+cdxWlzY+v+rFz+uHem5mYLgRqobkCRlOQNLmDsX+SnjyPFuGUc+wVMiRVVI5XbhBur83W2s4cqOZ223Va7aWfVheuWo58IxBYLnotxA2523NbfM5GyRZ7yPWXbTS/8l6/48vDE/eL1K63PHOQoDJ1UjbZXlTeGepgV0yX5FexT8gcnUc+S/NUTN6v0GDkcvMpgHXfeizShhbS7hCC72qXWjk+zqKd+SVFDVM2TUlSzhZLE7/o9urIxVa5FL9sF+ocR21lTAux2x7F8+KRPOjenU5q/+U2psU3rVjG6DQBoA0AaANAGgDQBoA0AaANAGgDQBoA0AaA+NAHMPn5h/wB9AapFXunzlksNPiLFFyFoqZGalkg2zgDv4Ykin3mJ3yP5gKl9Zf06A2QHbUyotyuW4rlHkZEFkxVKPLnyFMB5/wDxoD0aANAGgKaL5nrNXD33QATczeprK2wPcxZWcFRsv2VmwNJ7TcqyphbMMeZJkWCKZV6jNiPRHy7ovNi5+IcnBIpzq1dWXa64Nvn36R9VaKx2qyd6I51DIv5cuSJrQqme1c94mt+XIj6qq7xd8E4g+/zPqrNXv1YqmRGq62zrwQzK1E1qaRPNkd5uW3JGu1l5y3QLZwf87Su7fAMU8sfD5zhYGbvdFhCviVwjhOyzCwJoZqxlGkHtlj1TH5umiHJMOoWwAVI7EzaG3yGq0TXhLrQtWWyVj421tNGuaQOeqNjfFtyRF1sqd2ermqUr1SN0D4oreqWq0SXh12omOlsNW9FudHHt+pudwVdO3/hp/pTgT7LzNy3JvOIbhjHeY8EQ+8bH84vWcm4hrkNk/G2Ta6Q7OUmayoKUqziHinIigkMVcF2ZlAEzdYTpGL2HLkho75RuErViLAH81UM76S52anjuNBXwb2SSFHMkSF/XlvteJV2xybPNc83OlGxW294bbe6aRYauhjZV0lXFvXvj2Oa1V4csnazM+Bc2rvHvRWyxBeJTN+z6t3jIMfDPpi5YrkXdkaItAGHfvm7Rw0WdEZLc+gFjIgv2/mRObpIPsFHUust8n0i+TytddYKaaSvsFY6ri3NHU8r2QzsV25u2Ij1jSXV/LcuTV3qKSu03CXEej1tRVMje+egkWZurnG5yMemeqvBrK3Wy/CvBwEebDxE1KsBzGMY51IQTmMYRE4gxMUDH/SPIOXV84/lHXHX0bv8Ag4rXNVV01oVyqq75UpZmoq9SuyRE1vOVE2qpCdCfF9X6YPDUfLXp4XiGgDQBoA0AaANAGgDQBoA0AaANAGgDQCBbUH0k63PcRls8lZZ+0j884rQi2T+UevmMQ2V261pws1hmbk5k2iR1THWUTQKQh1jnVMAnOYRgOFJppMS4ia5znNZXUzWIqqqMRaCncqNTqzcqqpCsNPe7EV+RXOVErKbVRXKqN+40+xqLsairtVG5JnmvCo/up8TUrtyViDMlnntzL2qz9qr1XyE3cQ1vhSyUqnI2KEr2LU0a68wo4QP0xEjJOP6imFidspkEwftumS+M0B9Wrq7Ixcnh9Sr5FTvNiyXSbhDyJ28satI08YCEcykm+tXX2WzdqozfNXTDu+IM6IJCImByU4gRwjjrL1RgKvV4uo5DmaxaKvlyfiJkHyh7hj2wP4568Jjy8yPcI6eMHbh13q24VUWWYHAWKogQjc+gJZpEBWpzcWxmbNRbXXZihH8nr0itA21EL3kE1W8Bachzc8it4BSMbtzOI6IRWT6V3Cr2QOHUZh0AWBaANAGgIE3TUyo5C2251qF6rkTbarMYovZJWAnGib2NfFa1xw+bd5BX85JZNNVI5eR0lSEVTMU5CmDQ4pghqcOVzZGo9v1Wd2S/6mRuc136K1yI5FTaioioaXEdNT1lhrI5WNkYtNNm1yZouUbnIvpRURUVNqKiKm1CgmQnJub+DSQslOTMpNSS239gzWkZV6u+fLM2OQTMGLVZ0uInORFBJJFMDCPSmmQv5oaom+SySaAkVzlXfQNTPqa25NRrfQ1ERETqRETgQoe4yyzeTg9z3K538PVM3LmuSTq1qehrURqJ1IiIWf5W/BSuf2WaT/d/u3HayNJfRlrvZ6H3ISd3/mdf3RB4URKW0X6B2NPdLOfvPtfLRb0YKX2fuHu1ZnYN5r6fu+b9pTGbDfuHnf14X+DNrlX6N3i7FPrbT8NMRfQlxdV+mD3FHz16dl4ivxW9Ha5PZFLiaDzLV5rIJ7XIUYteiE5iSONviVzNZWveYsmx2niWqpDpOSeI+TqkMmr0HKIABhWO/XZzKymQYaK3C47lpLFzd07uTaKk15PwrSOkQh5l1CLMUlE5hOOeD4OVPDGflinfNrIi1XAU9AeOv8QPZnapUYSA3B0WRk0lnLZ63IeWRLEOWbMZB43n3DlqRKPURb/HqkfHbmSREqpwKmcphAZmTvlLhaU4yRLWmCjaA0r5bW5uLySat64hWjswkCTqssoYEQaigYFQWE3R2xA3PloBerjvq2hY/tErTrjn7H8FYIBKHXsLdzIOFGVbRn4xOaiFrJNNUjso8qzNZJ2BnjhACNlCLqdKRgNoCfGmSseSCFBdR14qkkzyr1DjN7Gz8ZIMcgELW3Fx7tOes1DpSJBimrmRA7Q6pDMkFHJRFIom0BAV+33bO8XX55jHIG4vF1XucQ5aMbMwkbEj5dR5CQIVWOjsjWVADxdbcuSHTO2bzzyOWcJqoqIkORZMTANcgug6QRctlknDZwkmu3cIKEWQXQWJ3ElkVUxEpimKICUwCICA8wHloD9dAGgDQBoA0AaAr52k/Sj4kfv+xR/LhWNV7hHlRiTvCl+XUxB8Mco7/wBtpvgKcsG1YROCvbJt6lfSfNruyZismKrbjg7I2G8fxrhFuytzbyYj6vTJa6qmJrb51JCoxWZNu8Lbl5eJUXHtmAkql7go6HqmX5O+WaMd2+nWu1rpUFGSYr2logxr8fIo1aNg24+LWAjlczdBXtG7gHIqJ+g3XoCF6bkLN56zeK9ZZ6/QeT6YjH7hq56b1aHhG9xpb4ikzasOeEinL1JRpF8zQxXIKkeoLnaOzEVBI5VQMtYMj5oc0Gnt42Ttx8p5kXPlpRjSIKNtDbElDh2qUlXacRF8rHkVbPHhY6IkHSoAvIJLzi6KaShCGRAdnG13aZIolXuzRk7iwn4tJy9hpBMyUjAzCBjMp6uySZgDk5j3qbhk45cyd5A/QYxOkwgbxoA0BFWdvvIZk91WQ/qi81psRcn67slT4LzWXriar7NP4TjnQL+LLwPuHR/xLW1QF55gU9ZF80Q58rOja7u93xLi1LK34KV1+yzSfq3HazNJfRlrvZ6Hw4Sf3/mdf3RB4URKW0X6B2NPdLOfvPtfLRb0YKX2fuHu1ZnYN5r6fu+b9pTGbDfuHnf14X+DNrlX6N3i7FPrbT8NMRfQlxdV+mD3FHz16dl4lLeyWv7gbTjfItb8m29o4AsW5berBWC0kuWQofOUXV5nOdojZl9DNkoFSKRk0VTrGROEp2CoAgoKpFgOQoGf2xSFt2bX7bNsSyREYmzDi260e+wu1/cTjhqzgrmeFxnGI2CTrObsYreKQIuqycF/tZXZRxGTL9I5pCFgVnCJVgMFX6mtasZ8ZaCr8GjLWKTy/mCJhWTNmgo+cSDzbbXkmzJr0l6us5zcilD1iJv79AbLn/J2Pn3BgkLO0tMIvBXbZzTKrU3Cb1sfz60Wuhsa1X6tGIFMJlZFxIKkZFYkAXAOupAyYHIcAA3Dh8xUejb98bCeYxorw+4WoxUojIIs1fBvYnbxUGb9s4FYBAOyoVQo8/UA9X9+gENhgax+WHUftdMgfF8Lv13dttrTeDMsapM74HDLt7vLsVjUA+ILEt8oHsRE02HyFGS8zQZgDdMgaAsg4fkPttc8PvELqDaUuRo83iosjnJ7b0ox4vKZNfR5/wCkIvmRaf6xNMDYfOU7GSTHmk6K5bnKRBMhCgYDhJLuF9q0qWCXkneDWmc82MNq7uWO5UWX22Mruu2xr5Yo7ETmikyFcJQo/aeVEaAj8mBHQFn2gDQBoA0AaANAV87SfpR8SP3/AGKP5cKxqvcI8qMSd4Uvy6mIPhjlHf8AttN8BTlg2rCJwIjkjO+R6/Zc1x8VjyvWpSsIrM8P2Q0TJOGVXscNj8lttQ5ceIifwrBuA+ZMHLc7TzJEwQ6IhKB1iBjK9mDIM3O5MsSqREo2mpS7lmkTEBxqj0I3FUTbiJSmTRfmMVVV29VKJCo8+RCNg5CXvHA27NmS7qNLwi1rrmTj5PLbcEbJKUapPbdbY5p6F+kjv0VjmL1sqh1LdJDuO6qKCQhyKJzAoUDVpfcPfcbjeS2FrDvqvUcZVRuztspX31fsMBkuXo57LFp5HhlHixG0XKLimyQ6ViHYygCydrqC9ZrHAyrbNWT4tGZvb2co72rM8uMceBjj0dWiJ1yzko+MMdzXLIm/UFeTI5erOjtnDHtLoAZLrbGT8QoB847y7kSbol1ut6vrymxsXXpKSTsE1hBaGqUMdnY1mhHcFIryyppsDoJkSRbk6FHAqpqpiKihUNAMHg59kyXpQTmUHjNaTnJJ7KQMelX0q5KwtSciBq/H2hm2evkRkxR5LPOyoVNFRXw4FEURUOB7s7feQzJ7qsh/VF5rTYi5P13ZKnwXmsvXE1X2afwnHOgX8WXgfcOj/iWtqgLzzAp6yL5ohz5WdG13d7viXFqWVvwUrr9lmk/VuO1maS+jLXez0Phwk/v/ADOv7og8KIlLaL9A7GnulnP3n2vlot6MFL7P3D3aszsG819P3fN+0pjNhv3Dzv68L/Bm1yr9G7xdin1tp+GmIvoS4uq/TB7ij569Oy8RRYnYjtWgb2rkiExgtE2le3zd8c+X3/JrWtu7XZXZ39jlHlFRmQg1vHrKqqvG6scZq6OocyyJxObmBk8M7JtrW3y6S2Q8Q4ghKlcZaNcwhZoZSzz56/X3smpNPq1RmdofPW9di13ap3C8ZAJRrBZXoMo3MKafSBgdre4LZVmy17gmG1DJuLb3cqxko59xUfQpQHM5FZKWYlr4u7ixW6VCrqIxfhCrlJ2DixURIcVEFAADV5vZvsMwjO2Dc9Z8Y0ChJY/kLDmKYstksVga4roc2RJSWsWS2WPJeQNU4iRKbvPlJVnENnhXiiz0q4PFlFTgKDkG48FnIGEbPxDbrkjG32A8rWptBW/OsXk7L1Ood3ubTnT0WM5C1KRYoOpD5KLE/djhXVKj0KicC6AtKqGKsIp1vDS9EotJjqlitoWcwc3q0KyhoCksZ+nu6uV5TY2MIig1TcQ0o8a8iJAUW7tX2eo/PQCA5UxBwn7XvBLgTJcTjQN1+aoZXLEnhJC13+tI5cjIop/H3C8YxrL9pWLAosm0UB6rNMHjiSQamI7B0ijyKBvmRs11PdZi3dDtG4du6vEeON0uHEYHEk/IQzFCYfbafFzLWHmpFCipppFFywhUpJtCCimeObTSTRJY5CoKJgApm2bYRxN8GZ6pMpuB4rFg3D7NMQydgyE0qs/jqKqOb7xLDT3UBEVTK14iyKt3lZjVHJ5gyKaniHLxo0IfoSII6AsPr/EF2X2rbddd31d3E49l9tOOpGdibtmFm9eqVOvSdafoxk0weq9jvd1Fdwgn0EROKgrpdrrBQoiBpmW+KTw98D4+w3lXL+6/E9Bx7uDgFrVhe1Tsq9SjciV1s3QdOJmvCg3Oc6BCOmwmOchOkVkyj7RuWgF8/wAvrwb/APmEbfv/ADcx/wDS0A4G1bfrtC3vp3NxtOzpUM5sMerRLW5SlJCVcxUE9nElHEWxcyL1sikZZVNFQ/bTMcxSl5nAvMOYDe6Ar52k/Sj4kfv+xR/LhWNV7hHlRiTvCl+XUxB8Mco7/wBtpvgKcsG1YROBVl7dZbNb8kwuMMU1OxVKMsLWp5Omp6bTg3NynVY1u1ssbEtCIqAr5awXTTWWe9CblQh2iPqJ3RAjEllw6XL2WGcZiIslMVyAn3FYdQNlBc+W7RTa0hD5LpzCloOwQK5j2qsTGgZy2N40xHhAABilNARNGPqy+wBMWysYixXPIVmz1iSWqbbKVyQeVpR6zRhCRjteTEH8HLRniQZPGvJFq5TSN2usnQAAbpPZQjaGrO4qlsU4oftUKxQKveaGvcjy1vyInP1hNRSs0mDnUVlZ1NszUUaNCO1e49Fq4IXtpk6wAyTWxQENaS5/+wPWk8av8jPK6S+Jz6z64xThSWTx19lT0SddTFBB27bJNTGZcpYI8COVRMQx0dAMwltrwgjHzUSSgsBiLARuSVh1JKdWiFytJj0haAjFrOjIIdp78qS8Oml21gAxOXSHICRqhSKxQ49xF1WOUjWLt8tJOElZGUkzKPXAAVZYXEssuoHPpD2QOBQ/IAevQGrZ2+8hmT3VZD+qLzWmxFyfruyVPgvNZeuJqvs0/hOOdAv4svA+4dH/ABLW1QF55gU9ZF80Q58rOja7u93xLi1LK34KV1+yzSfq3HazNJfRlrvZ6Hw4Sf3/AJnX90QeFESltF+gdjT3Szn7z7Xy0W9GCl9n7h7tWZ2Dea+n7vm/aUxmw37h539eF/gza5V+jd4uxT620/DTEX0JcXVfpg9xR89enZeIaANAfzD9p0Ruv2d533/8Y3aW7lbxC7Vt/uWMS7xdtrRDssMibU7Q/TtE3aWh2/PrWiHplV1TGROaMORhMpj4ZlIouALYd/W9qx/CEZea2fbFrhcaXw9MKYJW3P76dxbWAfMpGfloymuLvSNuzSOkgRIdyi9akTcNBVVSeSbd26EFWNfAXwFc9uImb4GRRBFJL1brCnL8WT1GLmqUKVX5vt+Xzm+f5/yaA/oo4J9WEMNgHqD7FWPOQB6gD+yLP5g0BzJ59TTH4WlsyOKaYn/oM3wesUyCcBBhYygYDiHPmACIAP5AEQ+YR0BRW2xLvHxrxJOL9xQ9hE8/sGY9iO8uSVyvtsTjHTiLz1truou5HJEUmnGD3HazQI8zpwxBMXPYIEnFH82YNkVwO5fZ/vuwZxHNkzXdDgSW8TWrdSLIystWfLtlbLje9R8Cb0nx9b26HqTesFFC8j8gSeNFGz9t1NXSJhA4/traSQfBGuIoAJJAH2Y9xpukEkwL1J5QrvbN08uXMvSXpH5y9JeXLkGgOlHh1bSNr+5/hW8N9PcTgHE2aiU3a/jc9ULkmjwNsCumlqw1LJ+UebIqdjv9lLugnyA/bIJuYlDQFOHHGxNszpMxiThZ8PbYttXkuIRvgMEAjOQuHqd4vbnhF6dRK0ZWlZBk1E8Y8XboPPAvDB1R0cylZf2Fm8d4gDpS4cOwTD3DW2m412uYfaIOEKrHFf3y8KMEGU5k/I8mmVW13ywCn1G7jpcOhqgZRQGMek0YJGFJsURAezQFfO0n6UfEj9/2KP5cKxqvcI8qMSd4Uvy6mIPhjlHf+203wFOWDasInAiOT3eNYrJl1sUUnmxBvCnrUlnuVxU+8LSWykQ1Tdxrm3NTG7zly1YFRPKhEB4gsUBQcd5QoIaA8cSx2mpFlRg4SWiLHhS/owb6yRzd1HXV9MXftkf21KwJnBWRjZpaYVB9JFP2nDozkDdKhChoDAU5DDLh/OQVhjMxTsPmyai6AzzldlUDM7s/x/JqqQNeRlI4U3Ddv3m5ix7940IaXV8T1L+pLuANtYsE48s7uxycjGuCzlilqpYPP27jtTcBP0hmiyrUxWJLp7rFZAiJQEUTAVUDrEVKZNdYhwNYbbZaC2sqEt5jbF6wyuDjIUbjJedVPjyNu7p0eRWnmsF0gP8Ataqr4jQyhmab5Q7kiAGHkADFaANARVnb7yGZPdVkP6ovNabEXJ+u7JU+C81l64mq+zT+E450C/iy8D7h0f8AEtbVAXnmBT1kXzRDnys6Nru73fEuLUsrfgpXX7LNJ+rcdrM0l9GWu9nofDhJ/f8Amdf3RB4URKW0X6B2NPdLOfvPtfLRb0YKX2fuHu1ZnYN5r6fu+b9pTGbDfuHnf14X+DNrlX6N3i7FPrbT8NMRfQlxdV+mD3FHz16dl4hoA0ByhfBooyNm5LjLQ0wwaSsRLcRLJcZKRkg3SdsJGPfQnhXrF61XASKJKpmMRRM4CU5TCUwCA6At7a7GNr+wTYJu6xNtVxhGYwpFkxpuKyHNx7Nw7kHMhZLNRZJw5O4kpI6i4t2qfbZxzUT9hgyRSatiJpE5aA5Ccd4EyDuC+BsKweNIN/ZZ+gZTuGXXsHFNln0o9rVCzS8eWZRiybFMoqZszVVdnKQoiCLdU/zFHQHXhw7uJTs13TbPMNZPomfcWtCxGL6bF3+sWC7VyBs+OLNB1tBlPV65Qko5TWZLN1UVORlC9ldHoct1VW6iapgKQcG5XpPEU+E9/wBILbBLtsobedmG0WYxtfMz1gVXtFkci2Px7BCCgLAUvhn5e7IiRJdqqok4Bm7USE6SXcMBOnAtAp+Jdx/ymADFNvCrhTFMACUxRJNgIGAdALhvjwZk/gG7p77xHdolOnbTw2t0Cx4Xf/tepRCdvE9onhWZR+bcbQqog0aNxcuznTN8UjGO3LiMP/U8oinGgIztOkGEt8EN4g8pFLLOIySyvuIfxy7hDwrhdi9yXXXDNVdr1H7ZzJmKJk+s3QIiXqNy56Avow9xA8QcND4P5tJ3QZacpuxgdqGK4bHNITcJoy2S8nSlO51KlRfUICUF1UxWfOAA3g45B066FDJFSUAWbgaY3xfSHOVuKJv53G4Ll+IVvmUC1SrKaylRwdYBwrLAg+qGJoZgq+Hy5dVokwM/ZkAgxzNpFwnQioxe98DproGacQZXXk22MMo4+yG4hE26swhSbfAWdWLSeCYGikgnDLrCiCglMBBP0gbpHlz5DoCTdAV87SfpR8SP3/Yo/lwrGq9wjyoxJ3hS/LqYg+GOUd/7bTfAU5YNqwicCYWqswpcmXXGNTzWjTFc0lVkbnRxrSU1KIyDqEKlPK0q0HOmhFSMnGJeIUaviSCvQCskxape0qAHktOCsYdqVS+yOtWpLHFsZXawum6LDpZY5nwb93GdqarlMVxCv0IgOkxxByisgR42VRcolU0BqcIhRkpFhUZbOLubxRhu0TV3YxZKBKshaOK4UbA1h7hlchDx7tlC+KKZum2SZPXBuwk+dPlAEhwGTT3I4aVj1ZELf0gk9g2HlysHYkJ9VxZm67uvma1xZoV8sm8SbOVG6yTc6ShWzgSnEEVOkD9j7icRAWH8NaFZNWdbSjtizhK/ZJyQIjCSSMPL+YR0S0WWaGbunCKCpHRETkVUIUQ5iGgMg2zvil1awpaVtbFnjzilXRIuylG0Y5s6TQHytaaTzhErFWQKkPUZkm4M5DkYop9RDAAG6QN1qdolbVCV+fjZeWo8ujA21gycFWcQMw4jUZdGPkCB9ooZuukpy/Qbl9sBgADUs8GITB+ZTqHTSTLinIhjqqqESSTIFReCY6qqggUpQD1iYRAAD1iPLWmxFyfruyVP6fkvNZe8ks1X2afh2J/acc5yRyKfBloExDdRfsEph1Bz5CJcmrFHlz1QF55gf+5F80Q58rOja/sD0/8AZcWq5W/BSuv2WaT9W47WZpL6Mtd7PQ+HCT+/8zr+6IPCiJS2i/QOxp7pZz959r5aLejBS+z9w92rM7BvNfT93zftKYzYb9w87+vC/wAGbXKv0bvF2KfW2n4aYi+hLi6r9MHuKPnr07LxDQBoBV9s2yvbXs9c5edbeMcIY+WzvkV9lfKJkJuxTAWa+SSIIPZw5Z9057AnAPWk27SPMRN0cx0Aw1vqkBe6pZqRamBZWsXGvzFXscWdVdAkjBT8epFSzE67UxFSAqgqomJkzlOXq5lMA8h0BCO1naTgHZdhKC267cqGhQcQVxzOO4qpDKzNhSScWN8eSmlF39kcO3KvfWUOY4KqmD1iHLQFbeZvg7HB9zpkGQybbtnlRhrPNSa0vPkx/PWyg1+bfuFe84Wc1esPW8eh1j9sVi3aFHmI8uoerQFmW2vanty2eY3Z4j2xYco+FseM1zu/R+lxRWfmD9T1KSc9KuDKvpJ2IciC7kHLlz2ykT7vbIQoAa9gvZlty22ZLz9l7DWPU6hkDc/cWt+zXOEm7BKDcLU0TUTRkhZzDpdBp/rlTCkzTQSE6hzdHMdAMBb6hVsgVWx0a8V6ItlNuEJJ1q01iwMG8pB2CAmmZ4+Wh5aOdlMku3cIKHSVTOUSnIYQENAJBUOFxsaoWz697C6hhBnB7U8kvrBI23Fja13dVvIOrK/byUmonYnUipKpfGtGwpiR8AkBAnrEeoTAR9nLg18O/cjjzBGJM0YKWumMdtFN9BcKURzkTJMfWqVBmSSbqqN42IlkCuHiiSCCJ3rvvuhRSIn3AIHLQCyf5tDwU/8Agrr/AP7Fy1/+7oB29lfC32McPKZv0/tCwbG4jlsnR8JFXV4ysdvn1JiPrq6zqIaj6UP3vaKko4WPyS6OoxxE3PQFgegK+dpP0o+JH7/sUfy4VjVe4R5UYk7wpfl1MQfDHKO/9tpvgKcsG1YROCvK3YfyUtuaZWuv32mMbQ4mGlsURkbSckxYsLRr9tHy1OicYiQyLOSZKCih6appOXKse4VhF1kE5BPwIEiXPBcfmSVzLa4LIAuW1yhK1TIcKldpBvDsbDj15Kx84ztCUJ3EVDovHAoOEPaWKKKzZcqRg5aA01pgHKqIZUqB0EfKcgNbo3ZW4cjTK1cjPSSMZptVkcYKN+02WSVQOCaqKh/Dm6ViFU6zlADLjtVuAWJssplK1SMenP48nG1xkLH/AKS65F1OuTUM+ocDJoxoMhjCupEki0VWa+KWM9lW7w4lBooAGiNMMXmsXyHZw9iYZHtlWhrYW0Is8xvqTkCdZWm3xkzVsi39mzQOZdw7axflroO2kyMqxIDAqbU3ZQA31xtqyCzsCN2jrmSaco5StWRi4tsMmuniwHc1HmCty7dowZ+NRkoh2CSgF8UswcFM7VFHxB0BRA/Gv4Hz1h52ytGPb7Wsm2OQgFK1boy/MVKqxku7POLalajzUIV24dP0XryQR+UEAVmz/wBs4eFTKICcblxy5xIs8TOymjMb/izZdi9+xW3h5lWjZinyucXya/iWm3fEEk7IkqtFuO3/AGjmm/xQogdmicQ5A8qq+PuGOLs62QbrDboHf1Cp1XMWZzV/x4tZEzTNOFM2qqK928Yxs1T35blju8PtUKTU9rp3J/E6rVdG6qdwpR06rkqsX8yRM0Xb+FqJND+9i9p7mbJG8GrYzWayygouDgYzc5kGLjU1cebbMTQiyJ2tOjiNPiTT6wJJlI16u4mp0oH5rqOjNojjmpXE0sWE7LExI2LF9cnRM4aSOnkY/c/11HNasztbXe/KBqrI+V8UYxvWfzRKmDrJHG1m5sZcqhG509vpG5fZJlsWZctiZ562TPOWR0Ttb7LLjbblsTf4cl7G4UfzePofD2OGK/Q5sVpkYpi1jyOzNEA5iRJJIF3ipS9BOsiRfjVkCH13lCXaw4I0HVdsmnVZaqibbKCPzpqiVqMTPUb+FjW68rkTJqZNTfOaiybSPV2rDGjt9FJKutJTMoqZq7ZZXMRjc8k6kRM3rlkmaJ5zmos47Z6xYqjsnx/W7VDO4GwsMSSXmEK+ApX7A7xs7foIPEw+0V7SpBUSH2kjiKZvaKOtpgC0XSy+ThT0tXA+nqo8PVm6QSf3I1kiqZGten4X6j2q5q7WKuq7aim5wvR1dFo3gimjdFK23y60bvObrNkciO/3OyVM06l2LwGobDfuHnf1oX+DNrkT6N1f6din1tp+GmIdoS4uqvTB7jh89enZeIaANAGgDQBoA0AaANAGgDQBoA0AaANAV57RXCCu6fiUpJLEUUb7gMTEcEKPMyJzbbauoUqgfkESiA/9B1XuEeVGJO8KX5fT/wDwg2GFauJMQbeCtpc/0+4U6lhmrCJycguT8v7obVvqtPF+om3y0WPaptbznGbeo/LiN9lIycebBakm5xdvGcxe3JKO8fYEi3l4rkeHly/KlYioEQbEVTMmU4Gl4a3qbmdtkjuJq20vIRdzsqnk/irXy/bRCYJsDU+1NhW822/IWMM7OLw2aNXcwSdfuUHRa4oss5tbCwsggOg7IxygZyxcUzc7jjDGbXVF3txG5fA9dnttTVPfebAkHjm3VvIWe/NGNh294yplubx1cn5ltKIRsjzfJdNKrUis1sK6sgzIZICU67xEt7+Hc44Hxpac3xW5LMUng+yVvLmCW2NoCMxtPZtd7PXe5TCiuFcs0FsqWxK2CZbIQryfSepVgyr9aEYplfRp1tAKbR93WZ4K8Zr3Dwu7p3kK15axNw/8dbmd5ae26Vg43YpU8kZLtDvKtfj8TLxpEhPXnLgzNgeYJJGrIv0H9k600VQKAyTDicbrGJNo0paNyi03gcu7/IWDo3JtFwIZrnff5jdCxRcDiDNNExjOx52C9YYOH76BvQRKkVIO1GPpfX+7FJdpUDrf0BrFyraluqdiqyE/P1NSwRD6JLZKo+8qskJ49EUDSUFJFAwt3afUJkVwDqSU5HL7QBrGrKdaqlkjSR8Wu1W7pEurIzP8THdTk6l6jHq6daulkj13xa7VbukTtWRmezWY78Lk6l6lFP2+7UsN8P3Ct3icB4+uVzk3ryYvdmUWlkbVl7Ltxdc1Cec22zLN/FulTj20TO3KTduB1D8y9Sgmi9sw/bcD2iodR00tRIutK9rNR1TUyJ5rEVysan6Jm1jc3v4XOVYvZsN2rA9oqPqNNLM9yvnkRFa+qqpepFe9WoqrwN1nI1uar1qqpjti2r583D7h5bebvyqxqrJ1mSVjsB7fH0hGzMZQo1i4FRhPSpYtZy0OokI9xsQFDHXeCpJOgBTwyKFHYT0Y4jxvjtcT4tiRrqR+rZ7Q5UfDStauccr2oqtcrF3zdba+f7ZybyMrzC+EMQYnxM6+Yhh3FYXatutrnMe2BrVXVlfqOexVThYma5uzlfvtVGyLub3RS2VMhF2vYDUWmXDiRNCX2fiFuQyL9M3J7TIl+n6k2zcAEZt8A9KZSnZkNz8RyovynNNWIMc3j+QcH7pUT1cv1S51NM7VdO/8y2wTJ/agjTNbpV55MjR1O1c90MPSBjarvlx/gVpzkc925VUka+e78VOx6eaxv+0y9SZxouesP3hDEjLEFNawgLkfzjsiDmwyaZRTQcPipdAN2KP5jZAPi0S/OJQ6je0YddL+TzoHs2grCC0zHpU3SuWOou9ciarZqhrNRkMDPy6SlZ9lTs4VbnI/fPVEtHBuFYMKWtIkXXmkydUScCOflkjWJ+GNib1icOW1dqky6v8AJcGgDQBoA0AaANAGgDQBoA0AaANAGgDQHgaxcYxcyDxlHMWbyWXSdSrpq0bt3Mm5RblaIuJBdIoGWORIhEynUExiplKQB6SgGvw2KNjnKjWorlzcqIiK5css3L1rls29R+Gxxsc5UaiK5c3KiIiuVEyzcvWuWzb1Hv1+z9nmKzaEbCyI1bkZimdIWhUUythSU5gomKAB09JuY8w5ch5jz0B5WcLDR7t6/YRMYxfSXY8xes2DVs7f+GT7TbxrlEoHV7ZfZJ1ibpL6i8g0Apu5fZDh3czE45QlnFsxjZ8P3SVyDi294gmS0mx0612CGXr1jkG6LZI7F2WQYuV27tJ80cEUBTuB0qgB9Ab9tu2t4Z2qYcxvg3EVY8FS8WQqcJV3E+6WstmBIr1xJqu5CzS/cdrrqOXbpYTmU5FMuoVIpE+RAAm8K9AAhJtQg4cGs13hmGwRjIEJbxJRK48zR6OlfuAIgfugbqAR6uegPr6OV4AhgCBhgCu8vR8PK2PKC5IeGDyb2PkvxfxfxPR7Hs/a+rQGa0B//9k=";
    }
}