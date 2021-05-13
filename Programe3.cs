using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using Op = DocumentFormat.OpenXml.CustomProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;

namespace GeneratedCode
{
    public class GeneratedClass
    {
        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath)
        {
            using(SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId3");
            GenerateThemePart1Content(themePart1);

            PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart1 = workbookPart1.AddNewPart<PivotTableCacheDefinitionPart>("rId2");
            GeneratePivotTableCacheDefinitionPart1Content(pivotTableCacheDefinitionPart1);

            PivotTableCacheRecordsPart pivotTableCacheRecordsPart1 = pivotTableCacheDefinitionPart1.AddNewPart<PivotTableCacheRecordsPart>("rId1");
            GeneratePivotTableCacheRecordsPart1Content(pivotTableCacheRecordsPart1);

            WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            GenerateWorksheetPart1Content(worksheetPart1);

            PivotTablePart pivotTablePart1 = worksheetPart1.AddNewPart<PivotTablePart>("rId1");
            GeneratePivotTablePart1Content(pivotTablePart1);

            pivotTablePart1.AddPart(pivotTableCacheDefinitionPart1, "rId1");

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId5");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId4");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            CustomFilePropertiesPart customFilePropertiesPart1 = document.AddNewPart<CustomFilePropertiesPart>("rId4");
            GenerateCustomFilePropertiesPart1Content(customFilePropertiesPart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
        private void GenerateExtendedFilePropertiesPart1Content(ExtendedFilePropertiesPart extendedFilePropertiesPart1)
        {
            Ap.Properties properties1 = new Ap.Properties();
            properties1.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application1 = new Ap.Application();
            application1.Text = "Microsoft Excel Online";
            Ap.Manager manager1 = new Ap.Manager();
            manager1.Text = "";
            Ap.Company company1 = new Ap.Company();
            company1.Text = "";
            Ap.HyperlinkBase hyperlinkBase1 = new Ap.HyperlinkBase();
            hyperlinkBase1.Text = "";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "16.0300";

            properties1.Append(application1);
            properties1.Append(manager1);
            properties1.Append(company1);
            properties1.Append(hyperlinkBase1);
            properties1.Append(applicationVersion1);

            extendedFilePropertiesPart1.Properties = properties1;
        }

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "x15 xr xr6 xr10 xr2" }  };
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbook1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            workbook1.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            workbook1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            workbook1.AddNamespaceDeclaration("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6");
            workbook1.AddNamespaceDeclaration("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10");
            workbook1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            FileVersion fileVersion1 = new FileVersion(){ ApplicationName = "xl", LastEdited = "7", LowestEdited = "7", BuildVersion = "23902" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties();

            OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xr:revisionPtr revIDLastSave=\"14\" documentId=\"11_E6100F9A48B216FFE929DBA94E7110680B17EAEF\" xr6:coauthVersionLast=\"46\" xr6:coauthVersionMax=\"46\" xr10:uidLastSave=\"{B318D50E-0BE0-4246-AC7A-A8EE6D5C10AA}\" xmlns:xr10=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision10\" xmlns:xr6=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision6\" xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\" />");

            BookViews bookViews1 = new BookViews();

            WorkbookView workbookView1 = new WorkbookView(){ XWindow = 0, YWindow = 0, WindowWidth = (UInt32Value)24225U, WindowHeight = (UInt32Value)12540U };
            workbookView1.SetAttribute(new OpenXmlAttribute("xr2", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2", "{00000000-000D-0000-FFFF-FFFF00000000}"));

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet(){ Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };

            sheets1.Append(sheet1);
            CalculationProperties calculationProperties1 = new CalculationProperties(){ CalculationId = (UInt32Value)191028U, CalculationCompleted = false };

            PivotCaches pivotCaches1 = new PivotCaches();
            PivotCache pivotCache1 = new PivotCache(){ CacheId = (UInt32Value)49U, Id = "rId2" };

            pivotCaches1.Append(pivotCache1);

            WorkbookExtensionList workbookExtensionList1 = new WorkbookExtensionList();

            WorkbookExtension workbookExtension1 = new WorkbookExtension(){ Uri = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" };
            workbookExtension1.AddNamespaceDeclaration("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures");

            OpenXmlUnknownElement openXmlUnknownElement2 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xcalcf:calcFeatures xmlns:xcalcf=\"http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures\"><xcalcf:feature name=\"microsoft.com:RD\" /><xcalcf:feature name=\"microsoft.com:Single\" /><xcalcf:feature name=\"microsoft.com:FV\" /><xcalcf:feature name=\"microsoft.com:CNMTM\" /><xcalcf:feature name=\"microsoft.com:LET_WF\" /></xcalcf:calcFeatures>");

            workbookExtension1.Append(openXmlUnknownElement2);

            workbookExtensionList1.Append(workbookExtension1);

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(openXmlUnknownElement1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);
            workbook1.Append(pivotCaches1);
            workbook1.Append(workbookExtensionList1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme(){ Name = "Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme(){ Name = "Office" };

            A.Dark1Color dark1Color1 = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor(){ Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color1.Append(systemColor1);

            A.Light1Color light1Color1 = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor(){ Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color1.Append(systemColor2);

            A.Dark2Color dark2Color1 = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex(){ Val = "44546A" };

            dark2Color1.Append(rgbColorModelHex1);

            A.Light2Color light2Color1 = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex(){ Val = "E7E6E6" };

            light2Color1.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color1 = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex(){ Val = "5B9BD5" };

            accent1Color1.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color1 = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex(){ Val = "ED7D31" };

            accent2Color1.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color1 = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex(){ Val = "A5A5A5" };

            accent3Color1.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color1 = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex(){ Val = "FFC000" };

            accent4Color1.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color1 = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex(){ Val = "4472C4" };

            accent5Color1.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color1 = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex(){ Val = "70AD47" };

            accent6Color1.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink1 = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex(){ Val = "0563C1" };

            hyperlink1.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor1 = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex(){ Val = "954F72" };

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

            A.FontScheme fontScheme1 = new A.FontScheme(){ Name = "Office" };

            A.MajorFont majorFont1 = new A.MajorFont();
            A.LatinFont latinFont1 = new A.LatinFont(){ Typeface = "Calibri Light" };
            A.EastAsianFont eastAsianFont1 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont1 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont1 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont2 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont3 = new A.SupplementalFont(){ Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont4 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont5 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont6 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont7 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont8 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont9 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont10 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont11 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "MoolBoran" };
            A.SupplementalFont supplementalFont12 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont13 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont14 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont15 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont16 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont17 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont18 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont19 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont20 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont21 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont22 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont23 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont24 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont25 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont26 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont27 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont28 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Times New Roman" };
            A.SupplementalFont supplementalFont29 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont30 = new A.SupplementalFont(){ Script = "Geor", Typeface = "Sylfaen" };

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
            A.LatinFont latinFont2 = new A.LatinFont(){ Typeface = "Calibri" };
            A.EastAsianFont eastAsianFont2 = new A.EastAsianFont(){ Typeface = "" };
            A.ComplexScriptFont complexScriptFont2 = new A.ComplexScriptFont(){ Typeface = "" };
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont(){ Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont(){ Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont(){ Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont(){ Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont(){ Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont(){ Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont(){ Script = "Thai", Typeface = "Tahoma" };
            A.SupplementalFont supplementalFont38 = new A.SupplementalFont(){ Script = "Ethi", Typeface = "Nyala" };
            A.SupplementalFont supplementalFont39 = new A.SupplementalFont(){ Script = "Beng", Typeface = "Vrinda" };
            A.SupplementalFont supplementalFont40 = new A.SupplementalFont(){ Script = "Gujr", Typeface = "Shruti" };
            A.SupplementalFont supplementalFont41 = new A.SupplementalFont(){ Script = "Khmr", Typeface = "DaunPenh" };
            A.SupplementalFont supplementalFont42 = new A.SupplementalFont(){ Script = "Knda", Typeface = "Tunga" };
            A.SupplementalFont supplementalFont43 = new A.SupplementalFont(){ Script = "Guru", Typeface = "Raavi" };
            A.SupplementalFont supplementalFont44 = new A.SupplementalFont(){ Script = "Cans", Typeface = "Euphemia" };
            A.SupplementalFont supplementalFont45 = new A.SupplementalFont(){ Script = "Cher", Typeface = "Plantagenet Cherokee" };
            A.SupplementalFont supplementalFont46 = new A.SupplementalFont(){ Script = "Yiii", Typeface = "Microsoft Yi Baiti" };
            A.SupplementalFont supplementalFont47 = new A.SupplementalFont(){ Script = "Tibt", Typeface = "Microsoft Himalaya" };
            A.SupplementalFont supplementalFont48 = new A.SupplementalFont(){ Script = "Thaa", Typeface = "MV Boli" };
            A.SupplementalFont supplementalFont49 = new A.SupplementalFont(){ Script = "Deva", Typeface = "Mangal" };
            A.SupplementalFont supplementalFont50 = new A.SupplementalFont(){ Script = "Telu", Typeface = "Gautami" };
            A.SupplementalFont supplementalFont51 = new A.SupplementalFont(){ Script = "Taml", Typeface = "Latha" };
            A.SupplementalFont supplementalFont52 = new A.SupplementalFont(){ Script = "Syrc", Typeface = "Estrangelo Edessa" };
            A.SupplementalFont supplementalFont53 = new A.SupplementalFont(){ Script = "Orya", Typeface = "Kalinga" };
            A.SupplementalFont supplementalFont54 = new A.SupplementalFont(){ Script = "Mlym", Typeface = "Kartika" };
            A.SupplementalFont supplementalFont55 = new A.SupplementalFont(){ Script = "Laoo", Typeface = "DokChampa" };
            A.SupplementalFont supplementalFont56 = new A.SupplementalFont(){ Script = "Sinh", Typeface = "Iskoola Pota" };
            A.SupplementalFont supplementalFont57 = new A.SupplementalFont(){ Script = "Mong", Typeface = "Mongolian Baiti" };
            A.SupplementalFont supplementalFont58 = new A.SupplementalFont(){ Script = "Viet", Typeface = "Arial" };
            A.SupplementalFont supplementalFont59 = new A.SupplementalFont(){ Script = "Uigh", Typeface = "Microsoft Uighur" };
            A.SupplementalFont supplementalFont60 = new A.SupplementalFont(){ Script = "Geor", Typeface = "Sylfaen" };

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

            A.FormatScheme formatScheme1 = new A.FormatScheme(){ Name = "Office" };

            A.FillStyleList fillStyleList1 = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation(){ Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation(){ Val = 105000 };
            A.Tint tint1 = new A.Tint(){ Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation(){ Val = 103000 };
            A.Tint tint2 = new A.Tint(){ Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation(){ Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation(){ Val = 109000 };
            A.Tint tint3 = new A.Tint(){ Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation(){ Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation(){ Val = 102000 };
            A.Tint tint4 = new A.Tint(){ Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation(){ Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation(){ Val = 100000 };
            A.Shade shade1 = new A.Shade(){ Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation(){ Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation(){ Val = 120000 };
            A.Shade shade2 = new A.Shade(){ Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList1.Append(solidFill1);
            fillStyleList1.Append(gradientFill1);
            fillStyleList1.Append(gradientFill2);

            A.LineStyleList lineStyleList1 = new A.LineStyleList();

            A.Outline outline1 = new A.Outline(){ Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter(){ Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline(){ Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter(){ Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline(){ Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash(){ Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter(){ Limit = 800000 };

            outline3.Append(solidFill4);
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

            A.OuterShadow outerShadow1 = new A.OuterShadow(){ BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex(){ Val = "000000" };
            A.Alpha alpha1 = new A.Alpha(){ Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList1.Append(effectStyle1);
            effectStyleList1.Append(effectStyle2);
            effectStyleList1.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList1 = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint(){ Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation(){ Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill(){ RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop(){ Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint(){ Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation(){ Val = 150000 };
            A.Shade shade3 = new A.Shade(){ Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation(){ Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop(){ Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint(){ Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation(){ Val = 130000 };
            A.Shade shade4 = new A.Shade(){ Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation(){ Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop(){ Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor(){ Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade(){ Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation(){ Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill(){ Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList1.Append(solidFill5);
            backgroundFillStyleList1.Append(solidFill6);
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

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        // Generates content of pivotTableCacheDefinitionPart1.
        private void GeneratePivotTableCacheDefinitionPart1Content(PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart1)
        {
            PivotCacheDefinition pivotCacheDefinition1 = new PivotCacheDefinition(){ Id = "rId1", RefreshedBy = "Excel Services", RefreshedDate = 44262.096334837966D, CreatedVersion = 7, RefreshedVersion = 7, MinRefreshableVersion = 3, RecordCount = (UInt32Value)38U, MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "xr" }  };
            pivotCacheDefinition1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            pivotCacheDefinition1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            pivotCacheDefinition1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            pivotCacheDefinition1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{87720BC6-432E-4D37-BB27-89B609BD257B}"));

            CacheSource cacheSource1 = new CacheSource(){ Type = SourceValues.Worksheet };
            WorksheetSource worksheetSource1 = new WorksheetSource(){ Reference = "A1:AA1048576", Sheet = "Sheet1" };

            cacheSource1.Append(worksheetSource1);

            CacheFields cacheFields1 = new CacheFields(){ Count = (UInt32Value)27U };

            CacheField cacheField1 = new CacheField(){ Name = "城市", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems1 = new SharedItems(){ ContainsBlank = true };

            cacheField1.Append(sharedItems1);

            CacheField cacheField2 = new CacheField(){ Name = "负责人", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems2 = new SharedItems(){ ContainsBlank = true };

            cacheField2.Append(sharedItems2);

            CacheField cacheField3 = new CacheField(){ Name = "物理店", NumberFormatId = (UInt32Value)0U };

            SharedItems sharedItems3 = new SharedItems(){ ContainsBlank = true, Count = (UInt32Value)9U };
            StringItem stringItem1 = new StringItem(){ Val = "赣州" };
            StringItem stringItem2 = new StringItem(){ Val = "云南旗舰店" };
            StringItem stringItem3 = new StringItem(){ Val = "江汉" };
            StringItem stringItem4 = new StringItem(){ Val = "海珠" };
            StringItem stringItem5 = new StringItem(){ Val = "海口" };
            StringItem stringItem6 = new StringItem(){ Val = "厦门" };
            StringItem stringItem7 = new StringItem(){ Val = "新生" };
            StringItem stringItem8 = new StringItem(){ Val = "横岗" };
            MissingItem missingItem1 = new MissingItem();

            sharedItems3.Append(stringItem1);
            sharedItems3.Append(stringItem2);
            sharedItems3.Append(stringItem3);
            sharedItems3.Append(stringItem4);
            sharedItems3.Append(stringItem5);
            sharedItems3.Append(stringItem6);
            sharedItems3.Append(stringItem7);
            sharedItems3.Append(stringItem8);
            sharedItems3.Append(missingItem1);

            cacheField3.Append(sharedItems3);

            CacheField cacheField4 = new CacheField(){ Name = "门店ID", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems4 = new SharedItems(){ ContainsBlank = true };

            cacheField4.Append(sharedItems4);

            CacheField cacheField5 = new CacheField(){ Name = "门店", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems5 = new SharedItems(){ ContainsBlank = true };

            cacheField5.Append(sharedItems5);

            CacheField cacheField6 = new CacheField(){ Name = "平台", NumberFormatId = (UInt32Value)0U };

            SharedItems sharedItems6 = new SharedItems(){ ContainsBlank = true, Count = (UInt32Value)3U };
            StringItem stringItem9 = new StringItem(){ Val = "美团" };
            StringItem stringItem10 = new StringItem(){ Val = "饿了么" };
            MissingItem missingItem2 = new MissingItem();

            sharedItems6.Append(stringItem9);
            sharedItems6.Append(stringItem10);
            sharedItems6.Append(missingItem2);

            cacheField6.Append(sharedItems6);

            CacheField cacheField7 = new CacheField(){ Name = "三方配送", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems7 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, ContainsInteger = true, MinValue = 0D, MaxValue = 0D };

            cacheField7.Append(sharedItems7);

            CacheField cacheField8 = new CacheField(){ Name = "单价", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems8 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 8.6199999999999992D, MaxValue = 18.690000000000001D };

            cacheField8.Append(sharedItems8);

            CacheField cacheField9 = new CacheField(){ Name = "订单", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems9 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, ContainsInteger = true, MinValue = 4D, MaxValue = 82D };

            cacheField9.Append(sharedItems9);

            CacheField cacheField10 = new CacheField(){ Name = "收入", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems10 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 74.760000000000005D, MaxValue = 1260.67D };

            cacheField10.Append(sharedItems10);

            CacheField cacheField11 = new CacheField(){ Name = "平均收入", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems11 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 266.79000000000002D, MaxValue = 1610.425D };

            cacheField11.Append(sharedItems11);

            CacheField cacheField12 = new CacheField(){ Name = "总收入", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems12 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 1067.1600000000001D, MaxValue = 8024.5D };

            cacheField12.Append(sharedItems12);

            CacheField cacheField13 = new CacheField(){ Name = "成本", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems13 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 33.72D, MaxValue = 557.13D };

            cacheField13.Append(sharedItems13);

            CacheField cacheField14 = new CacheField(){ Name = "平均成本", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems14 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 117.83D, MaxValue = 736.74D };

            cacheField14.Append(sharedItems14);

            CacheField cacheField15 = new CacheField(){ Name = "总成本", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems15 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 471.32D, MaxValue = 2946.96D };

            cacheField15.Append(sharedItems15);

            CacheField cacheField16 = new CacheField(){ Name = "成本比例", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems16 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 0.318D, MaxValue = 0.72450000000000003D };

            cacheField16.Append(sharedItems16);

            CacheField cacheField17 = new CacheField(){ Name = "总成本比例", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems17 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 0.36122700000000002D, MaxValue = 0.49230299999999999D };

            cacheField17.Append(sharedItems17);

            CacheField cacheField18 = new CacheField(){ Name = "推广", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems18 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 0D, MaxValue = 32D };

            cacheField18.Append(sharedItems18);

            CacheField cacheField19 = new CacheField(){ Name = "平均推广", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems19 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 7.94D, MaxValue = 45.322499999999998D };

            cacheField19.Append(sharedItems19);

            CacheField cacheField20 = new CacheField(){ Name = "总推广", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems20 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 28.9D, MaxValue = 181.29D };

            cacheField20.Append(sharedItems20);

            CacheField cacheField21 = new CacheField(){ Name = "推广比例", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems21 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 0D, MaxValue = 6.4199999999999993E-2D };

            cacheField21.Append(sharedItems21);

            CacheField cacheField22 = new CacheField(){ Name = "总推广比例", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems22 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 1.7405E-2D, MaxValue = 8.8037000000000004E-2D };

            cacheField22.Append(sharedItems22);

            CacheField cacheField23 = new CacheField(){ Name = "比30天", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems23 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 0.48286602000000001D, MaxValue = 3.3815369099999999D };

            cacheField23.Append(sharedItems23);

            CacheField cacheField24 = new CacheField(){ Name = "比上天", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems24 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = -0.43385080999999998D, MaxValue = 1.0893506500000001D };

            cacheField24.Append(sharedItems24);

            CacheField cacheField25 = new CacheField(){ Name = "比上周", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems25 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = -0.50734760999999995D, MaxValue = 5.8615359500000004D };

            cacheField25.Append(sharedItems25);

            CacheField cacheField26 = new CacheField(){ Name = "比上周(3)", NumberFormatId = (UInt32Value)0U };
            SharedItems sharedItems26 = new SharedItems(){ ContainsString = false, ContainsBlank = true, ContainsNumber = true, MinValue = 0.77675992000000005D, MaxValue = 19.01719915D };

            cacheField26.Append(sharedItems26);

            CacheField cacheField27 = new CacheField(){ Name = "日期", NumberFormatId = (UInt32Value)0U };

            SharedItems sharedItems27 = new SharedItems(){ ContainsBlank = true, Count = (UInt32Value)2U };
            StringItem stringItem11 = new StringItem(){ Val = "20210302" };
            MissingItem missingItem3 = new MissingItem();

            sharedItems27.Append(stringItem11);
            sharedItems27.Append(missingItem3);

            cacheField27.Append(sharedItems27);

            cacheFields1.Append(cacheField1);
            cacheFields1.Append(cacheField2);
            cacheFields1.Append(cacheField3);
            cacheFields1.Append(cacheField4);
            cacheFields1.Append(cacheField5);
            cacheFields1.Append(cacheField6);
            cacheFields1.Append(cacheField7);
            cacheFields1.Append(cacheField8);
            cacheFields1.Append(cacheField9);
            cacheFields1.Append(cacheField10);
            cacheFields1.Append(cacheField11);
            cacheFields1.Append(cacheField12);
            cacheFields1.Append(cacheField13);
            cacheFields1.Append(cacheField14);
            cacheFields1.Append(cacheField15);
            cacheFields1.Append(cacheField16);
            cacheFields1.Append(cacheField17);
            cacheFields1.Append(cacheField18);
            cacheFields1.Append(cacheField19);
            cacheFields1.Append(cacheField20);
            cacheFields1.Append(cacheField21);
            cacheFields1.Append(cacheField22);
            cacheFields1.Append(cacheField23);
            cacheFields1.Append(cacheField24);
            cacheFields1.Append(cacheField25);
            cacheFields1.Append(cacheField26);
            cacheFields1.Append(cacheField27);

            PivotCacheDefinitionExtensionList pivotCacheDefinitionExtensionList1 = new PivotCacheDefinitionExtensionList();

            PivotCacheDefinitionExtension pivotCacheDefinitionExtension1 = new PivotCacheDefinitionExtension(){ Uri = "{725AE2AE-9491-48be-B2B4-4EB974FC3084}" };
            pivotCacheDefinitionExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.PivotCacheDefinition pivotCacheDefinition2 = new X14.PivotCacheDefinition();

            pivotCacheDefinitionExtension1.Append(pivotCacheDefinition2);

            pivotCacheDefinitionExtensionList1.Append(pivotCacheDefinitionExtension1);

            pivotCacheDefinition1.Append(cacheSource1);
            pivotCacheDefinition1.Append(cacheFields1);
            pivotCacheDefinition1.Append(pivotCacheDefinitionExtensionList1);

            pivotTableCacheDefinitionPart1.PivotCacheDefinition = pivotCacheDefinition1;
        }

        // Generates content of pivotTableCacheRecordsPart1.
        private void GeneratePivotTableCacheRecordsPart1Content(PivotTableCacheRecordsPart pivotTableCacheRecordsPart1)
        {
            PivotCacheRecords pivotCacheRecords1 = new PivotCacheRecords(){ Count = (UInt32Value)38U, MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "xr" }  };
            pivotCacheRecords1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            pivotCacheRecords1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            pivotCacheRecords1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            PivotCacheRecord pivotCacheRecord1 = new PivotCacheRecord();
            StringItem stringItem12 = new StringItem(){ Val = "赣州" };
            StringItem stringItem13 = new StringItem(){ Val = "小伍" };
            FieldItem fieldItem1 = new FieldItem(){ Val = (UInt32Value)0U };
            StringItem stringItem14 = new StringItem(){ Val = "11027801" };
            StringItem stringItem15 = new StringItem(){ Val = "金牌手抓饼•奶茶•小吃(赣州店)" };
            FieldItem fieldItem2 = new FieldItem(){ Val = (UInt32Value)0U };
            NumberItem numberItem1 = new NumberItem(){ Val = 0D };
            NumberItem numberItem2 = new NumberItem(){ Val = 12.5D };
            NumberItem numberItem3 = new NumberItem(){ Val = 17D };
            NumberItem numberItem4 = new NumberItem(){ Val = 212.47D };
            NumberItem numberItem5 = new NumberItem(){ Val = 1604.9D };
            NumberItem numberItem6 = new NumberItem(){ Val = 8024.5D };
            NumberItem numberItem7 = new NumberItem(){ Val = 67.56D };
            NumberItem numberItem8 = new NumberItem(){ Val = 579.73400000000004D };
            NumberItem numberItem9 = new NumberItem(){ Val = 2898.67D };
            NumberItem numberItem10 = new NumberItem(){ Val = 0.318D };
            NumberItem numberItem11 = new NumberItem(){ Val = 0.36122700000000002D };
            NumberItem numberItem12 = new NumberItem(){ Val = 2.99D };
            NumberItem numberItem13 = new NumberItem(){ Val = 36.177999999999997D };
            NumberItem numberItem14 = new NumberItem(){ Val = 180.89D };
            NumberItem numberItem15 = new NumberItem(){ Val = 1.41E-2D };
            NumberItem numberItem16 = new NumberItem(){ Val = 2.2542E-2D };
            NumberItem numberItem17 = new NumberItem(){ Val = 1.4769126699999999D };
            NumberItem numberItem18 = new NumberItem(){ Val = 0.13130290999999999D };
            MissingItem missingItem4 = new MissingItem();
            MissingItem missingItem5 = new MissingItem();
            FieldItem fieldItem3 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord1.Append(stringItem12);
            pivotCacheRecord1.Append(stringItem13);
            pivotCacheRecord1.Append(fieldItem1);
            pivotCacheRecord1.Append(stringItem14);
            pivotCacheRecord1.Append(stringItem15);
            pivotCacheRecord1.Append(fieldItem2);
            pivotCacheRecord1.Append(numberItem1);
            pivotCacheRecord1.Append(numberItem2);
            pivotCacheRecord1.Append(numberItem3);
            pivotCacheRecord1.Append(numberItem4);
            pivotCacheRecord1.Append(numberItem5);
            pivotCacheRecord1.Append(numberItem6);
            pivotCacheRecord1.Append(numberItem7);
            pivotCacheRecord1.Append(numberItem8);
            pivotCacheRecord1.Append(numberItem9);
            pivotCacheRecord1.Append(numberItem10);
            pivotCacheRecord1.Append(numberItem11);
            pivotCacheRecord1.Append(numberItem12);
            pivotCacheRecord1.Append(numberItem13);
            pivotCacheRecord1.Append(numberItem14);
            pivotCacheRecord1.Append(numberItem15);
            pivotCacheRecord1.Append(numberItem16);
            pivotCacheRecord1.Append(numberItem17);
            pivotCacheRecord1.Append(numberItem18);
            pivotCacheRecord1.Append(missingItem4);
            pivotCacheRecord1.Append(missingItem5);
            pivotCacheRecord1.Append(fieldItem3);

            PivotCacheRecord pivotCacheRecord2 = new PivotCacheRecord();
            StringItem stringItem16 = new StringItem(){ Val = "云南" };
            StringItem stringItem17 = new StringItem(){ Val = "小伍" };
            FieldItem fieldItem4 = new FieldItem(){ Val = (UInt32Value)1U };
            StringItem stringItem18 = new StringItem(){ Val = "2072014729" };
            StringItem stringItem19 = new StringItem(){ Val = "苏姐牛奶甜品世家(昭通店)" };
            FieldItem fieldItem5 = new FieldItem(){ Val = (UInt32Value)1U };
            NumberItem numberItem19 = new NumberItem(){ Val = 0D };
            NumberItem numberItem20 = new NumberItem(){ Val = 18.690000000000001D };
            NumberItem numberItem21 = new NumberItem(){ Val = 4D };
            NumberItem numberItem22 = new NumberItem(){ Val = 74.760000000000005D };
            NumberItem numberItem23 = new NumberItem(){ Val = 266.79000000000002D };
            NumberItem numberItem24 = new NumberItem(){ Val = 1067.1600000000001D };
            NumberItem numberItem25 = new NumberItem(){ Val = 33.72D };
            NumberItem numberItem26 = new NumberItem(){ Val = 117.83D };
            NumberItem numberItem27 = new NumberItem(){ Val = 471.32D };
            NumberItem numberItem28 = new NumberItem(){ Val = 0.45100000000000001D };
            NumberItem numberItem29 = new NumberItem(){ Val = 0.441658D };
            NumberItem numberItem30 = new NumberItem(){ Val = 0D };
            NumberItem numberItem31 = new NumberItem(){ Val = 23.487500000000001D };
            NumberItem numberItem32 = new NumberItem(){ Val = 93.95D };
            NumberItem numberItem33 = new NumberItem(){ Val = 0D };
            NumberItem numberItem34 = new NumberItem(){ Val = 8.8037000000000004E-2D };
            NumberItem numberItem35 = new NumberItem(){ Val = 0.48286602000000001D };
            NumberItem numberItem36 = new NumberItem(){ Val = -0.43385080999999998D };
            NumberItem numberItem37 = new NumberItem(){ Val = -0.50734760999999995D };
            NumberItem numberItem38 = new NumberItem(){ Val = 0.77675992000000005D };
            FieldItem fieldItem6 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord2.Append(stringItem16);
            pivotCacheRecord2.Append(stringItem17);
            pivotCacheRecord2.Append(fieldItem4);
            pivotCacheRecord2.Append(stringItem18);
            pivotCacheRecord2.Append(stringItem19);
            pivotCacheRecord2.Append(fieldItem5);
            pivotCacheRecord2.Append(numberItem19);
            pivotCacheRecord2.Append(numberItem20);
            pivotCacheRecord2.Append(numberItem21);
            pivotCacheRecord2.Append(numberItem22);
            pivotCacheRecord2.Append(numberItem23);
            pivotCacheRecord2.Append(numberItem24);
            pivotCacheRecord2.Append(numberItem25);
            pivotCacheRecord2.Append(numberItem26);
            pivotCacheRecord2.Append(numberItem27);
            pivotCacheRecord2.Append(numberItem28);
            pivotCacheRecord2.Append(numberItem29);
            pivotCacheRecord2.Append(numberItem30);
            pivotCacheRecord2.Append(numberItem31);
            pivotCacheRecord2.Append(numberItem32);
            pivotCacheRecord2.Append(numberItem33);
            pivotCacheRecord2.Append(numberItem34);
            pivotCacheRecord2.Append(numberItem35);
            pivotCacheRecord2.Append(numberItem36);
            pivotCacheRecord2.Append(numberItem37);
            pivotCacheRecord2.Append(numberItem38);
            pivotCacheRecord2.Append(fieldItem6);

            PivotCacheRecord pivotCacheRecord3 = new PivotCacheRecord();
            StringItem stringItem20 = new StringItem(){ Val = "武汉" };
            StringItem stringItem21 = new StringItem(){ Val = "邓信庭" };
            FieldItem fieldItem7 = new FieldItem(){ Val = (UInt32Value)2U };
            StringItem stringItem22 = new StringItem(){ Val = "11057742" };
            StringItem stringItem23 = new StringItem(){ Val = "古御贡茶•手抓饼•小吃（江汉店）" };
            FieldItem fieldItem8 = new FieldItem(){ Val = (UInt32Value)0U };
            NumberItem numberItem39 = new NumberItem(){ Val = 0D };
            NumberItem numberItem40 = new NumberItem(){ Val = 12.28D };
            NumberItem numberItem41 = new NumberItem(){ Val = 58D };
            NumberItem numberItem42 = new NumberItem(){ Val = 712.48D };
            NumberItem numberItem43 = new NumberItem(){ Val = 830.2D };
            NumberItem numberItem44 = new NumberItem(){ Val = 1660.4D };
            NumberItem numberItem45 = new NumberItem(){ Val = 389.25D };
            NumberItem numberItem46 = new NumberItem(){ Val = 408.71D };
            NumberItem numberItem47 = new NumberItem(){ Val = 817.42D };
            NumberItem numberItem48 = new NumberItem(){ Val = 0.54630000000000001D };
            NumberItem numberItem49 = new NumberItem(){ Val = 0.49230299999999999D };
            NumberItem numberItem50 = new NumberItem(){ Val = 0D };
            NumberItem numberItem51 = new NumberItem(){ Val = 14.45D };
            NumberItem numberItem52 = new NumberItem(){ Val = 28.9D };
            NumberItem numberItem53 = new NumberItem(){ Val = 0D };
            NumberItem numberItem54 = new NumberItem(){ Val = 1.7405E-2D };
            NumberItem numberItem55 = new NumberItem(){ Val = 1.60369742D };
            NumberItem numberItem56 = new NumberItem(){ Val = -7.2327900000000001E-2D };
            MissingItem missingItem6 = new MissingItem();
            MissingItem missingItem7 = new MissingItem();
            FieldItem fieldItem9 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord3.Append(stringItem20);
            pivotCacheRecord3.Append(stringItem21);
            pivotCacheRecord3.Append(fieldItem7);
            pivotCacheRecord3.Append(stringItem22);
            pivotCacheRecord3.Append(stringItem23);
            pivotCacheRecord3.Append(fieldItem8);
            pivotCacheRecord3.Append(numberItem39);
            pivotCacheRecord3.Append(numberItem40);
            pivotCacheRecord3.Append(numberItem41);
            pivotCacheRecord3.Append(numberItem42);
            pivotCacheRecord3.Append(numberItem43);
            pivotCacheRecord3.Append(numberItem44);
            pivotCacheRecord3.Append(numberItem45);
            pivotCacheRecord3.Append(numberItem46);
            pivotCacheRecord3.Append(numberItem47);
            pivotCacheRecord3.Append(numberItem48);
            pivotCacheRecord3.Append(numberItem49);
            pivotCacheRecord3.Append(numberItem50);
            pivotCacheRecord3.Append(numberItem51);
            pivotCacheRecord3.Append(numberItem52);
            pivotCacheRecord3.Append(numberItem53);
            pivotCacheRecord3.Append(numberItem54);
            pivotCacheRecord3.Append(numberItem55);
            pivotCacheRecord3.Append(numberItem56);
            pivotCacheRecord3.Append(missingItem6);
            pivotCacheRecord3.Append(missingItem7);
            pivotCacheRecord3.Append(fieldItem9);

            PivotCacheRecord pivotCacheRecord4 = new PivotCacheRecord();
            StringItem stringItem24 = new StringItem(){ Val = "广州" };
            StringItem stringItem25 = new StringItem(){ Val = "郑秀娟" };
            FieldItem fieldItem10 = new FieldItem(){ Val = (UInt32Value)3U };
            StringItem stringItem26 = new StringItem(){ Val = "2077906251" };
            StringItem stringItem27 = new StringItem(){ Val = "贡茶(海珠店)" };
            FieldItem fieldItem11 = new FieldItem(){ Val = (UInt32Value)1U };
            NumberItem numberItem57 = new NumberItem(){ Val = 0D };
            NumberItem numberItem58 = new NumberItem(){ Val = 10.78D };
            NumberItem numberItem59 = new NumberItem(){ Val = 42D };
            NumberItem numberItem60 = new NumberItem(){ Val = 452.84D };
            NumberItem numberItem61 = new NumberItem(){ Val = 798.62D };
            NumberItem numberItem62 = new NumberItem(){ Val = 3194.48D };
            NumberItem numberItem63 = new NumberItem(){ Val = 231.77D };
            NumberItem numberItem64 = new NumberItem(){ Val = 374.93D };
            NumberItem numberItem65 = new NumberItem(){ Val = 1499.72D };
            NumberItem numberItem66 = new NumberItem(){ Val = 0.51180000000000003D };
            NumberItem numberItem67 = new NumberItem(){ Val = 0.469472D };
            NumberItem numberItem68 = new NumberItem(){ Val = 18.899999999999999D };
            NumberItem numberItem69 = new NumberItem(){ Val = 33.207500000000003D };
            NumberItem numberItem70 = new NumberItem(){ Val = 132.83000000000001D };
            NumberItem numberItem71 = new NumberItem(){ Val = 4.1700000000000001E-2D };
            NumberItem numberItem72 = new NumberItem(){ Val = 4.1581E-2D };
            NumberItem numberItem73 = new NumberItem(){ Val = 1.50967355D };
            NumberItem numberItem74 = new NumberItem(){ Val = 0.15724105999999999D };
            MissingItem missingItem8 = new MissingItem();
            MissingItem missingItem9 = new MissingItem();
            FieldItem fieldItem12 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord4.Append(stringItem24);
            pivotCacheRecord4.Append(stringItem25);
            pivotCacheRecord4.Append(fieldItem10);
            pivotCacheRecord4.Append(stringItem26);
            pivotCacheRecord4.Append(stringItem27);
            pivotCacheRecord4.Append(fieldItem11);
            pivotCacheRecord4.Append(numberItem57);
            pivotCacheRecord4.Append(numberItem58);
            pivotCacheRecord4.Append(numberItem59);
            pivotCacheRecord4.Append(numberItem60);
            pivotCacheRecord4.Append(numberItem61);
            pivotCacheRecord4.Append(numberItem62);
            pivotCacheRecord4.Append(numberItem63);
            pivotCacheRecord4.Append(numberItem64);
            pivotCacheRecord4.Append(numberItem65);
            pivotCacheRecord4.Append(numberItem66);
            pivotCacheRecord4.Append(numberItem67);
            pivotCacheRecord4.Append(numberItem68);
            pivotCacheRecord4.Append(numberItem69);
            pivotCacheRecord4.Append(numberItem70);
            pivotCacheRecord4.Append(numberItem71);
            pivotCacheRecord4.Append(numberItem72);
            pivotCacheRecord4.Append(numberItem73);
            pivotCacheRecord4.Append(numberItem74);
            pivotCacheRecord4.Append(missingItem8);
            pivotCacheRecord4.Append(missingItem9);
            pivotCacheRecord4.Append(fieldItem12);

            PivotCacheRecord pivotCacheRecord5 = new PivotCacheRecord();
            StringItem stringItem28 = new StringItem(){ Val = "海南" };
            StringItem stringItem29 = new StringItem(){ Val = "于松民" };
            FieldItem fieldItem13 = new FieldItem(){ Val = (UInt32Value)4U };
            StringItem stringItem30 = new StringItem(){ Val = "10854598" };
            StringItem stringItem31 = new StringItem(){ Val = "喜三德甜品·手工芋圆（海口店）" };
            FieldItem fieldItem14 = new FieldItem(){ Val = (UInt32Value)0U };
            NumberItem numberItem75 = new NumberItem(){ Val = 0D };
            NumberItem numberItem76 = new NumberItem(){ Val = 15.37D };
            NumberItem numberItem77 = new NumberItem(){ Val = 82D };
            NumberItem numberItem78 = new NumberItem(){ Val = 1260.67D };
            NumberItem numberItem79 = new NumberItem(){ Val = 1016.7725D };
            NumberItem numberItem80 = new NumberItem(){ Val = 4067.09D };
            NumberItem numberItem81 = new NumberItem(){ Val = 557.13D };
            NumberItem numberItem82 = new NumberItem(){ Val = 469.32749999999999D };
            NumberItem numberItem83 = new NumberItem(){ Val = 1877.31D };
            NumberItem numberItem84 = new NumberItem(){ Val = 0.44190000000000002D };
            NumberItem numberItem85 = new NumberItem(){ Val = 0.461586D };
            NumberItem numberItem86 = new NumberItem(){ Val = 1.46D };
            NumberItem numberItem87 = new NumberItem(){ Val = 21.852499999999999D };
            NumberItem numberItem88 = new NumberItem(){ Val = 87.41D };
            NumberItem numberItem89 = new NumberItem(){ Val = 1.1999999999999999E-3D };
            NumberItem numberItem90 = new NumberItem(){ Val = 2.1492000000000001E-2D };
            NumberItem numberItem91 = new NumberItem(){ Val = 1.1638756400000001D };
            NumberItem numberItem92 = new NumberItem(){ Val = 8.0052769999999995E-2D };
            NumberItem numberItem93 = new NumberItem(){ Val = 5.8615359500000004D };
            NumberItem numberItem94 = new NumberItem(){ Val = 19.01719915D };
            FieldItem fieldItem15 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord5.Append(stringItem28);
            pivotCacheRecord5.Append(stringItem29);
            pivotCacheRecord5.Append(fieldItem13);
            pivotCacheRecord5.Append(stringItem30);
            pivotCacheRecord5.Append(stringItem31);
            pivotCacheRecord5.Append(fieldItem14);
            pivotCacheRecord5.Append(numberItem75);
            pivotCacheRecord5.Append(numberItem76);
            pivotCacheRecord5.Append(numberItem77);
            pivotCacheRecord5.Append(numberItem78);
            pivotCacheRecord5.Append(numberItem79);
            pivotCacheRecord5.Append(numberItem80);
            pivotCacheRecord5.Append(numberItem81);
            pivotCacheRecord5.Append(numberItem82);
            pivotCacheRecord5.Append(numberItem83);
            pivotCacheRecord5.Append(numberItem84);
            pivotCacheRecord5.Append(numberItem85);
            pivotCacheRecord5.Append(numberItem86);
            pivotCacheRecord5.Append(numberItem87);
            pivotCacheRecord5.Append(numberItem88);
            pivotCacheRecord5.Append(numberItem89);
            pivotCacheRecord5.Append(numberItem90);
            pivotCacheRecord5.Append(numberItem91);
            pivotCacheRecord5.Append(numberItem92);
            pivotCacheRecord5.Append(numberItem93);
            pivotCacheRecord5.Append(numberItem94);
            pivotCacheRecord5.Append(fieldItem15);

            PivotCacheRecord pivotCacheRecord6 = new PivotCacheRecord();
            StringItem stringItem32 = new StringItem(){ Val = "厦门" };
            StringItem stringItem33 = new StringItem(){ Val = "于松民" };
            FieldItem fieldItem16 = new FieldItem(){ Val = (UInt32Value)5U };
            StringItem stringItem34 = new StringItem(){ Val = "2077997044" };
            StringItem stringItem35 = new StringItem(){ Val = "喜三德甜品●手工芋圆(厦门店)" };
            FieldItem fieldItem17 = new FieldItem(){ Val = (UInt32Value)1U };
            NumberItem numberItem95 = new NumberItem(){ Val = 0D };
            NumberItem numberItem96 = new NumberItem(){ Val = 12.73D };
            NumberItem numberItem97 = new NumberItem(){ Val = 64D };
            NumberItem numberItem98 = new NumberItem(){ Val = 814.99D };
            NumberItem numberItem99 = new NumberItem(){ Val = 1182.8900000000001D };
            NumberItem numberItem100 = new NumberItem(){ Val = 4731.5600000000004D };
            NumberItem numberItem101 = new NumberItem(){ Val = 399.62D };
            NumberItem numberItem102 = new NumberItem(){ Val = 500.58749999999998D };
            NumberItem numberItem103 = new NumberItem(){ Val = 2002.35D };
            NumberItem numberItem104 = new NumberItem(){ Val = 0.49030000000000001D };
            NumberItem numberItem105 = new NumberItem(){ Val = 0.42319000000000001D };
            NumberItem numberItem106 = new NumberItem(){ Val = 32D };
            NumberItem numberItem107 = new NumberItem(){ Val = 29.675000000000001D };
            NumberItem numberItem108 = new NumberItem(){ Val = 118.7D };
            NumberItem numberItem109 = new NumberItem(){ Val = 3.9300000000000002E-2D };
            NumberItem numberItem110 = new NumberItem(){ Val = 2.5087000000000002E-2D };
            NumberItem numberItem111 = new NumberItem(){ Val = 1.8507105500000001D };
            NumberItem numberItem112 = new NumberItem(){ Val = 0.2040747D };
            MissingItem missingItem10 = new MissingItem();
            MissingItem missingItem11 = new MissingItem();
            FieldItem fieldItem18 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord6.Append(stringItem32);
            pivotCacheRecord6.Append(stringItem33);
            pivotCacheRecord6.Append(fieldItem16);
            pivotCacheRecord6.Append(stringItem34);
            pivotCacheRecord6.Append(stringItem35);
            pivotCacheRecord6.Append(fieldItem17);
            pivotCacheRecord6.Append(numberItem95);
            pivotCacheRecord6.Append(numberItem96);
            pivotCacheRecord6.Append(numberItem97);
            pivotCacheRecord6.Append(numberItem98);
            pivotCacheRecord6.Append(numberItem99);
            pivotCacheRecord6.Append(numberItem100);
            pivotCacheRecord6.Append(numberItem101);
            pivotCacheRecord6.Append(numberItem102);
            pivotCacheRecord6.Append(numberItem103);
            pivotCacheRecord6.Append(numberItem104);
            pivotCacheRecord6.Append(numberItem105);
            pivotCacheRecord6.Append(numberItem106);
            pivotCacheRecord6.Append(numberItem107);
            pivotCacheRecord6.Append(numberItem108);
            pivotCacheRecord6.Append(numberItem109);
            pivotCacheRecord6.Append(numberItem110);
            pivotCacheRecord6.Append(numberItem111);
            pivotCacheRecord6.Append(numberItem112);
            pivotCacheRecord6.Append(missingItem10);
            pivotCacheRecord6.Append(missingItem11);
            pivotCacheRecord6.Append(fieldItem18);

            PivotCacheRecord pivotCacheRecord7 = new PivotCacheRecord();
            StringItem stringItem36 = new StringItem(){ Val = "深圳" };
            StringItem stringItem37 = new StringItem(){ Val = "刘文君" };
            FieldItem fieldItem19 = new FieldItem(){ Val = (UInt32Value)6U };
            StringItem stringItem38 = new StringItem(){ Val = "501849656" };
            StringItem stringItem39 = new StringItem(){ Val = "贡茶(龙岗店)" };
            FieldItem fieldItem20 = new FieldItem(){ Val = (UInt32Value)1U };
            NumberItem numberItem113 = new NumberItem(){ Val = 0D };
            NumberItem numberItem114 = new NumberItem(){ Val = 8.6199999999999992D };
            NumberItem numberItem115 = new NumberItem(){ Val = 14D };
            NumberItem numberItem116 = new NumberItem(){ Val = 120.66D };
            NumberItem numberItem117 = new NumberItem(){ Val = 307.91199999999998D };
            NumberItem numberItem118 = new NumberItem(){ Val = 1539.56D };
            NumberItem numberItem119 = new NumberItem(){ Val = 87.42D };
            NumberItem numberItem120 = new NumberItem(){ Val = 140.934D };
            NumberItem numberItem121 = new NumberItem(){ Val = 704.67D };
            NumberItem numberItem122 = new NumberItem(){ Val = 0.72450000000000003D };
            NumberItem numberItem123 = new NumberItem(){ Val = 0.45770899999999998D };
            NumberItem numberItem124 = new NumberItem(){ Val = 4.0999999999999996D };
            NumberItem numberItem125 = new NumberItem(){ Val = 7.94D };
            NumberItem numberItem126 = new NumberItem(){ Val = 39.700000000000003D };
            NumberItem numberItem127 = new NumberItem(){ Val = 3.4000000000000002E-2D };
            NumberItem numberItem128 = new NumberItem(){ Val = 2.5787000000000001E-2D };
            NumberItem numberItem129 = new NumberItem(){ Val = 3.3815369099999999D };
            NumberItem numberItem130 = new NumberItem(){ Val = 1.0893506500000001D };
            MissingItem missingItem12 = new MissingItem();
            MissingItem missingItem13 = new MissingItem();
            FieldItem fieldItem21 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord7.Append(stringItem36);
            pivotCacheRecord7.Append(stringItem37);
            pivotCacheRecord7.Append(fieldItem19);
            pivotCacheRecord7.Append(stringItem38);
            pivotCacheRecord7.Append(stringItem39);
            pivotCacheRecord7.Append(fieldItem20);
            pivotCacheRecord7.Append(numberItem113);
            pivotCacheRecord7.Append(numberItem114);
            pivotCacheRecord7.Append(numberItem115);
            pivotCacheRecord7.Append(numberItem116);
            pivotCacheRecord7.Append(numberItem117);
            pivotCacheRecord7.Append(numberItem118);
            pivotCacheRecord7.Append(numberItem119);
            pivotCacheRecord7.Append(numberItem120);
            pivotCacheRecord7.Append(numberItem121);
            pivotCacheRecord7.Append(numberItem122);
            pivotCacheRecord7.Append(numberItem123);
            pivotCacheRecord7.Append(numberItem124);
            pivotCacheRecord7.Append(numberItem125);
            pivotCacheRecord7.Append(numberItem126);
            pivotCacheRecord7.Append(numberItem127);
            pivotCacheRecord7.Append(numberItem128);
            pivotCacheRecord7.Append(numberItem129);
            pivotCacheRecord7.Append(numberItem130);
            pivotCacheRecord7.Append(missingItem12);
            pivotCacheRecord7.Append(missingItem13);
            pivotCacheRecord7.Append(fieldItem21);

            PivotCacheRecord pivotCacheRecord8 = new PivotCacheRecord();
            StringItem stringItem40 = new StringItem(){ Val = "深圳" };
            StringItem stringItem41 = new StringItem(){ Val = "刘文君" };
            FieldItem fieldItem22 = new FieldItem(){ Val = (UInt32Value)7U };
            StringItem stringItem42 = new StringItem(){ Val = "501809918" };
            StringItem stringItem43 = new StringItem(){ Val = "苏姐牛奶甜品世家(华西二路店)" };
            FieldItem fieldItem23 = new FieldItem(){ Val = (UInt32Value)1U };
            NumberItem numberItem131 = new NumberItem(){ Val = 0D };
            NumberItem numberItem132 = new NumberItem(){ Val = 13.04D };
            NumberItem numberItem133 = new NumberItem(){ Val = 35D };
            NumberItem numberItem134 = new NumberItem(){ Val = 456.49D };
            NumberItem numberItem135 = new NumberItem(){ Val = 1610.425D };
            NumberItem numberItem136 = new NumberItem(){ Val = 6441.7D };
            NumberItem numberItem137 = new NumberItem(){ Val = 237.24D };
            NumberItem numberItem138 = new NumberItem(){ Val = 736.74D };
            NumberItem numberItem139 = new NumberItem(){ Val = 2946.96D };
            NumberItem numberItem140 = new NumberItem(){ Val = 0.51970000000000005D };
            NumberItem numberItem141 = new NumberItem(){ Val = 0.457482D };
            NumberItem numberItem142 = new NumberItem(){ Val = 29.3D };
            NumberItem numberItem143 = new NumberItem(){ Val = 45.322499999999998D };
            NumberItem numberItem144 = new NumberItem(){ Val = 181.29D };
            NumberItem numberItem145 = new NumberItem(){ Val = 6.4199999999999993E-2D };
            NumberItem numberItem146 = new NumberItem(){ Val = 2.8143000000000001E-2D };
            NumberItem numberItem147 = new NumberItem(){ Val = 1.0274903799999999D };
            NumberItem numberItem148 = new NumberItem(){ Val = 5.0174839999999998E-2D };
            MissingItem missingItem14 = new MissingItem();
            MissingItem missingItem15 = new MissingItem();
            FieldItem fieldItem24 = new FieldItem(){ Val = (UInt32Value)0U };

            pivotCacheRecord8.Append(stringItem40);
            pivotCacheRecord8.Append(stringItem41);
            pivotCacheRecord8.Append(fieldItem22);
            pivotCacheRecord8.Append(stringItem42);
            pivotCacheRecord8.Append(stringItem43);
            pivotCacheRecord8.Append(fieldItem23);
            pivotCacheRecord8.Append(numberItem131);
            pivotCacheRecord8.Append(numberItem132);
            pivotCacheRecord8.Append(numberItem133);
            pivotCacheRecord8.Append(numberItem134);
            pivotCacheRecord8.Append(numberItem135);
            pivotCacheRecord8.Append(numberItem136);
            pivotCacheRecord8.Append(numberItem137);
            pivotCacheRecord8.Append(numberItem138);
            pivotCacheRecord8.Append(numberItem139);
            pivotCacheRecord8.Append(numberItem140);
            pivotCacheRecord8.Append(numberItem141);
            pivotCacheRecord8.Append(numberItem142);
            pivotCacheRecord8.Append(numberItem143);
            pivotCacheRecord8.Append(numberItem144);
            pivotCacheRecord8.Append(numberItem145);
            pivotCacheRecord8.Append(numberItem146);
            pivotCacheRecord8.Append(numberItem147);
            pivotCacheRecord8.Append(numberItem148);
            pivotCacheRecord8.Append(missingItem14);
            pivotCacheRecord8.Append(missingItem15);
            pivotCacheRecord8.Append(fieldItem24);

            PivotCacheRecord pivotCacheRecord9 = new PivotCacheRecord();
            MissingItem missingItem16 = new MissingItem();
            MissingItem missingItem17 = new MissingItem();
            FieldItem fieldItem25 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem18 = new MissingItem();
            MissingItem missingItem19 = new MissingItem();
            FieldItem fieldItem26 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem20 = new MissingItem();
            MissingItem missingItem21 = new MissingItem();
            MissingItem missingItem22 = new MissingItem();
            MissingItem missingItem23 = new MissingItem();
            MissingItem missingItem24 = new MissingItem();
            MissingItem missingItem25 = new MissingItem();
            MissingItem missingItem26 = new MissingItem();
            MissingItem missingItem27 = new MissingItem();
            MissingItem missingItem28 = new MissingItem();
            MissingItem missingItem29 = new MissingItem();
            MissingItem missingItem30 = new MissingItem();
            MissingItem missingItem31 = new MissingItem();
            MissingItem missingItem32 = new MissingItem();
            MissingItem missingItem33 = new MissingItem();
            MissingItem missingItem34 = new MissingItem();
            MissingItem missingItem35 = new MissingItem();
            MissingItem missingItem36 = new MissingItem();
            MissingItem missingItem37 = new MissingItem();
            MissingItem missingItem38 = new MissingItem();
            MissingItem missingItem39 = new MissingItem();
            FieldItem fieldItem27 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord9.Append(missingItem16);
            pivotCacheRecord9.Append(missingItem17);
            pivotCacheRecord9.Append(fieldItem25);
            pivotCacheRecord9.Append(missingItem18);
            pivotCacheRecord9.Append(missingItem19);
            pivotCacheRecord9.Append(fieldItem26);
            pivotCacheRecord9.Append(missingItem20);
            pivotCacheRecord9.Append(missingItem21);
            pivotCacheRecord9.Append(missingItem22);
            pivotCacheRecord9.Append(missingItem23);
            pivotCacheRecord9.Append(missingItem24);
            pivotCacheRecord9.Append(missingItem25);
            pivotCacheRecord9.Append(missingItem26);
            pivotCacheRecord9.Append(missingItem27);
            pivotCacheRecord9.Append(missingItem28);
            pivotCacheRecord9.Append(missingItem29);
            pivotCacheRecord9.Append(missingItem30);
            pivotCacheRecord9.Append(missingItem31);
            pivotCacheRecord9.Append(missingItem32);
            pivotCacheRecord9.Append(missingItem33);
            pivotCacheRecord9.Append(missingItem34);
            pivotCacheRecord9.Append(missingItem35);
            pivotCacheRecord9.Append(missingItem36);
            pivotCacheRecord9.Append(missingItem37);
            pivotCacheRecord9.Append(missingItem38);
            pivotCacheRecord9.Append(missingItem39);
            pivotCacheRecord9.Append(fieldItem27);

            PivotCacheRecord pivotCacheRecord10 = new PivotCacheRecord();
            MissingItem missingItem40 = new MissingItem();
            MissingItem missingItem41 = new MissingItem();
            FieldItem fieldItem28 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem42 = new MissingItem();
            MissingItem missingItem43 = new MissingItem();
            FieldItem fieldItem29 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem44 = new MissingItem();
            MissingItem missingItem45 = new MissingItem();
            MissingItem missingItem46 = new MissingItem();
            MissingItem missingItem47 = new MissingItem();
            MissingItem missingItem48 = new MissingItem();
            MissingItem missingItem49 = new MissingItem();
            MissingItem missingItem50 = new MissingItem();
            MissingItem missingItem51 = new MissingItem();
            MissingItem missingItem52 = new MissingItem();
            MissingItem missingItem53 = new MissingItem();
            MissingItem missingItem54 = new MissingItem();
            MissingItem missingItem55 = new MissingItem();
            MissingItem missingItem56 = new MissingItem();
            MissingItem missingItem57 = new MissingItem();
            MissingItem missingItem58 = new MissingItem();
            MissingItem missingItem59 = new MissingItem();
            MissingItem missingItem60 = new MissingItem();
            MissingItem missingItem61 = new MissingItem();
            MissingItem missingItem62 = new MissingItem();
            MissingItem missingItem63 = new MissingItem();
            FieldItem fieldItem30 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord10.Append(missingItem40);
            pivotCacheRecord10.Append(missingItem41);
            pivotCacheRecord10.Append(fieldItem28);
            pivotCacheRecord10.Append(missingItem42);
            pivotCacheRecord10.Append(missingItem43);
            pivotCacheRecord10.Append(fieldItem29);
            pivotCacheRecord10.Append(missingItem44);
            pivotCacheRecord10.Append(missingItem45);
            pivotCacheRecord10.Append(missingItem46);
            pivotCacheRecord10.Append(missingItem47);
            pivotCacheRecord10.Append(missingItem48);
            pivotCacheRecord10.Append(missingItem49);
            pivotCacheRecord10.Append(missingItem50);
            pivotCacheRecord10.Append(missingItem51);
            pivotCacheRecord10.Append(missingItem52);
            pivotCacheRecord10.Append(missingItem53);
            pivotCacheRecord10.Append(missingItem54);
            pivotCacheRecord10.Append(missingItem55);
            pivotCacheRecord10.Append(missingItem56);
            pivotCacheRecord10.Append(missingItem57);
            pivotCacheRecord10.Append(missingItem58);
            pivotCacheRecord10.Append(missingItem59);
            pivotCacheRecord10.Append(missingItem60);
            pivotCacheRecord10.Append(missingItem61);
            pivotCacheRecord10.Append(missingItem62);
            pivotCacheRecord10.Append(missingItem63);
            pivotCacheRecord10.Append(fieldItem30);

            PivotCacheRecord pivotCacheRecord11 = new PivotCacheRecord();
            MissingItem missingItem64 = new MissingItem();
            MissingItem missingItem65 = new MissingItem();
            FieldItem fieldItem31 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem66 = new MissingItem();
            MissingItem missingItem67 = new MissingItem();
            FieldItem fieldItem32 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem68 = new MissingItem();
            MissingItem missingItem69 = new MissingItem();
            MissingItem missingItem70 = new MissingItem();
            MissingItem missingItem71 = new MissingItem();
            MissingItem missingItem72 = new MissingItem();
            MissingItem missingItem73 = new MissingItem();
            MissingItem missingItem74 = new MissingItem();
            MissingItem missingItem75 = new MissingItem();
            MissingItem missingItem76 = new MissingItem();
            MissingItem missingItem77 = new MissingItem();
            MissingItem missingItem78 = new MissingItem();
            MissingItem missingItem79 = new MissingItem();
            MissingItem missingItem80 = new MissingItem();
            MissingItem missingItem81 = new MissingItem();
            MissingItem missingItem82 = new MissingItem();
            MissingItem missingItem83 = new MissingItem();
            MissingItem missingItem84 = new MissingItem();
            MissingItem missingItem85 = new MissingItem();
            MissingItem missingItem86 = new MissingItem();
            MissingItem missingItem87 = new MissingItem();
            FieldItem fieldItem33 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord11.Append(missingItem64);
            pivotCacheRecord11.Append(missingItem65);
            pivotCacheRecord11.Append(fieldItem31);
            pivotCacheRecord11.Append(missingItem66);
            pivotCacheRecord11.Append(missingItem67);
            pivotCacheRecord11.Append(fieldItem32);
            pivotCacheRecord11.Append(missingItem68);
            pivotCacheRecord11.Append(missingItem69);
            pivotCacheRecord11.Append(missingItem70);
            pivotCacheRecord11.Append(missingItem71);
            pivotCacheRecord11.Append(missingItem72);
            pivotCacheRecord11.Append(missingItem73);
            pivotCacheRecord11.Append(missingItem74);
            pivotCacheRecord11.Append(missingItem75);
            pivotCacheRecord11.Append(missingItem76);
            pivotCacheRecord11.Append(missingItem77);
            pivotCacheRecord11.Append(missingItem78);
            pivotCacheRecord11.Append(missingItem79);
            pivotCacheRecord11.Append(missingItem80);
            pivotCacheRecord11.Append(missingItem81);
            pivotCacheRecord11.Append(missingItem82);
            pivotCacheRecord11.Append(missingItem83);
            pivotCacheRecord11.Append(missingItem84);
            pivotCacheRecord11.Append(missingItem85);
            pivotCacheRecord11.Append(missingItem86);
            pivotCacheRecord11.Append(missingItem87);
            pivotCacheRecord11.Append(fieldItem33);

            PivotCacheRecord pivotCacheRecord12 = new PivotCacheRecord();
            MissingItem missingItem88 = new MissingItem();
            MissingItem missingItem89 = new MissingItem();
            FieldItem fieldItem34 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem90 = new MissingItem();
            MissingItem missingItem91 = new MissingItem();
            FieldItem fieldItem35 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem92 = new MissingItem();
            MissingItem missingItem93 = new MissingItem();
            MissingItem missingItem94 = new MissingItem();
            MissingItem missingItem95 = new MissingItem();
            MissingItem missingItem96 = new MissingItem();
            MissingItem missingItem97 = new MissingItem();
            MissingItem missingItem98 = new MissingItem();
            MissingItem missingItem99 = new MissingItem();
            MissingItem missingItem100 = new MissingItem();
            MissingItem missingItem101 = new MissingItem();
            MissingItem missingItem102 = new MissingItem();
            MissingItem missingItem103 = new MissingItem();
            MissingItem missingItem104 = new MissingItem();
            MissingItem missingItem105 = new MissingItem();
            MissingItem missingItem106 = new MissingItem();
            MissingItem missingItem107 = new MissingItem();
            MissingItem missingItem108 = new MissingItem();
            MissingItem missingItem109 = new MissingItem();
            MissingItem missingItem110 = new MissingItem();
            MissingItem missingItem111 = new MissingItem();
            FieldItem fieldItem36 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord12.Append(missingItem88);
            pivotCacheRecord12.Append(missingItem89);
            pivotCacheRecord12.Append(fieldItem34);
            pivotCacheRecord12.Append(missingItem90);
            pivotCacheRecord12.Append(missingItem91);
            pivotCacheRecord12.Append(fieldItem35);
            pivotCacheRecord12.Append(missingItem92);
            pivotCacheRecord12.Append(missingItem93);
            pivotCacheRecord12.Append(missingItem94);
            pivotCacheRecord12.Append(missingItem95);
            pivotCacheRecord12.Append(missingItem96);
            pivotCacheRecord12.Append(missingItem97);
            pivotCacheRecord12.Append(missingItem98);
            pivotCacheRecord12.Append(missingItem99);
            pivotCacheRecord12.Append(missingItem100);
            pivotCacheRecord12.Append(missingItem101);
            pivotCacheRecord12.Append(missingItem102);
            pivotCacheRecord12.Append(missingItem103);
            pivotCacheRecord12.Append(missingItem104);
            pivotCacheRecord12.Append(missingItem105);
            pivotCacheRecord12.Append(missingItem106);
            pivotCacheRecord12.Append(missingItem107);
            pivotCacheRecord12.Append(missingItem108);
            pivotCacheRecord12.Append(missingItem109);
            pivotCacheRecord12.Append(missingItem110);
            pivotCacheRecord12.Append(missingItem111);
            pivotCacheRecord12.Append(fieldItem36);

            PivotCacheRecord pivotCacheRecord13 = new PivotCacheRecord();
            MissingItem missingItem112 = new MissingItem();
            MissingItem missingItem113 = new MissingItem();
            FieldItem fieldItem37 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem114 = new MissingItem();
            MissingItem missingItem115 = new MissingItem();
            FieldItem fieldItem38 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem116 = new MissingItem();
            MissingItem missingItem117 = new MissingItem();
            MissingItem missingItem118 = new MissingItem();
            MissingItem missingItem119 = new MissingItem();
            MissingItem missingItem120 = new MissingItem();
            MissingItem missingItem121 = new MissingItem();
            MissingItem missingItem122 = new MissingItem();
            MissingItem missingItem123 = new MissingItem();
            MissingItem missingItem124 = new MissingItem();
            MissingItem missingItem125 = new MissingItem();
            MissingItem missingItem126 = new MissingItem();
            MissingItem missingItem127 = new MissingItem();
            MissingItem missingItem128 = new MissingItem();
            MissingItem missingItem129 = new MissingItem();
            MissingItem missingItem130 = new MissingItem();
            MissingItem missingItem131 = new MissingItem();
            MissingItem missingItem132 = new MissingItem();
            MissingItem missingItem133 = new MissingItem();
            MissingItem missingItem134 = new MissingItem();
            MissingItem missingItem135 = new MissingItem();
            FieldItem fieldItem39 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord13.Append(missingItem112);
            pivotCacheRecord13.Append(missingItem113);
            pivotCacheRecord13.Append(fieldItem37);
            pivotCacheRecord13.Append(missingItem114);
            pivotCacheRecord13.Append(missingItem115);
            pivotCacheRecord13.Append(fieldItem38);
            pivotCacheRecord13.Append(missingItem116);
            pivotCacheRecord13.Append(missingItem117);
            pivotCacheRecord13.Append(missingItem118);
            pivotCacheRecord13.Append(missingItem119);
            pivotCacheRecord13.Append(missingItem120);
            pivotCacheRecord13.Append(missingItem121);
            pivotCacheRecord13.Append(missingItem122);
            pivotCacheRecord13.Append(missingItem123);
            pivotCacheRecord13.Append(missingItem124);
            pivotCacheRecord13.Append(missingItem125);
            pivotCacheRecord13.Append(missingItem126);
            pivotCacheRecord13.Append(missingItem127);
            pivotCacheRecord13.Append(missingItem128);
            pivotCacheRecord13.Append(missingItem129);
            pivotCacheRecord13.Append(missingItem130);
            pivotCacheRecord13.Append(missingItem131);
            pivotCacheRecord13.Append(missingItem132);
            pivotCacheRecord13.Append(missingItem133);
            pivotCacheRecord13.Append(missingItem134);
            pivotCacheRecord13.Append(missingItem135);
            pivotCacheRecord13.Append(fieldItem39);

            PivotCacheRecord pivotCacheRecord14 = new PivotCacheRecord();
            MissingItem missingItem136 = new MissingItem();
            MissingItem missingItem137 = new MissingItem();
            FieldItem fieldItem40 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem138 = new MissingItem();
            MissingItem missingItem139 = new MissingItem();
            FieldItem fieldItem41 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem140 = new MissingItem();
            MissingItem missingItem141 = new MissingItem();
            MissingItem missingItem142 = new MissingItem();
            MissingItem missingItem143 = new MissingItem();
            MissingItem missingItem144 = new MissingItem();
            MissingItem missingItem145 = new MissingItem();
            MissingItem missingItem146 = new MissingItem();
            MissingItem missingItem147 = new MissingItem();
            MissingItem missingItem148 = new MissingItem();
            MissingItem missingItem149 = new MissingItem();
            MissingItem missingItem150 = new MissingItem();
            MissingItem missingItem151 = new MissingItem();
            MissingItem missingItem152 = new MissingItem();
            MissingItem missingItem153 = new MissingItem();
            MissingItem missingItem154 = new MissingItem();
            MissingItem missingItem155 = new MissingItem();
            MissingItem missingItem156 = new MissingItem();
            MissingItem missingItem157 = new MissingItem();
            MissingItem missingItem158 = new MissingItem();
            MissingItem missingItem159 = new MissingItem();
            FieldItem fieldItem42 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord14.Append(missingItem136);
            pivotCacheRecord14.Append(missingItem137);
            pivotCacheRecord14.Append(fieldItem40);
            pivotCacheRecord14.Append(missingItem138);
            pivotCacheRecord14.Append(missingItem139);
            pivotCacheRecord14.Append(fieldItem41);
            pivotCacheRecord14.Append(missingItem140);
            pivotCacheRecord14.Append(missingItem141);
            pivotCacheRecord14.Append(missingItem142);
            pivotCacheRecord14.Append(missingItem143);
            pivotCacheRecord14.Append(missingItem144);
            pivotCacheRecord14.Append(missingItem145);
            pivotCacheRecord14.Append(missingItem146);
            pivotCacheRecord14.Append(missingItem147);
            pivotCacheRecord14.Append(missingItem148);
            pivotCacheRecord14.Append(missingItem149);
            pivotCacheRecord14.Append(missingItem150);
            pivotCacheRecord14.Append(missingItem151);
            pivotCacheRecord14.Append(missingItem152);
            pivotCacheRecord14.Append(missingItem153);
            pivotCacheRecord14.Append(missingItem154);
            pivotCacheRecord14.Append(missingItem155);
            pivotCacheRecord14.Append(missingItem156);
            pivotCacheRecord14.Append(missingItem157);
            pivotCacheRecord14.Append(missingItem158);
            pivotCacheRecord14.Append(missingItem159);
            pivotCacheRecord14.Append(fieldItem42);

            PivotCacheRecord pivotCacheRecord15 = new PivotCacheRecord();
            MissingItem missingItem160 = new MissingItem();
            MissingItem missingItem161 = new MissingItem();
            FieldItem fieldItem43 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem162 = new MissingItem();
            MissingItem missingItem163 = new MissingItem();
            FieldItem fieldItem44 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem164 = new MissingItem();
            MissingItem missingItem165 = new MissingItem();
            MissingItem missingItem166 = new MissingItem();
            MissingItem missingItem167 = new MissingItem();
            MissingItem missingItem168 = new MissingItem();
            MissingItem missingItem169 = new MissingItem();
            MissingItem missingItem170 = new MissingItem();
            MissingItem missingItem171 = new MissingItem();
            MissingItem missingItem172 = new MissingItem();
            MissingItem missingItem173 = new MissingItem();
            MissingItem missingItem174 = new MissingItem();
            MissingItem missingItem175 = new MissingItem();
            MissingItem missingItem176 = new MissingItem();
            MissingItem missingItem177 = new MissingItem();
            MissingItem missingItem178 = new MissingItem();
            MissingItem missingItem179 = new MissingItem();
            MissingItem missingItem180 = new MissingItem();
            MissingItem missingItem181 = new MissingItem();
            MissingItem missingItem182 = new MissingItem();
            MissingItem missingItem183 = new MissingItem();
            FieldItem fieldItem45 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord15.Append(missingItem160);
            pivotCacheRecord15.Append(missingItem161);
            pivotCacheRecord15.Append(fieldItem43);
            pivotCacheRecord15.Append(missingItem162);
            pivotCacheRecord15.Append(missingItem163);
            pivotCacheRecord15.Append(fieldItem44);
            pivotCacheRecord15.Append(missingItem164);
            pivotCacheRecord15.Append(missingItem165);
            pivotCacheRecord15.Append(missingItem166);
            pivotCacheRecord15.Append(missingItem167);
            pivotCacheRecord15.Append(missingItem168);
            pivotCacheRecord15.Append(missingItem169);
            pivotCacheRecord15.Append(missingItem170);
            pivotCacheRecord15.Append(missingItem171);
            pivotCacheRecord15.Append(missingItem172);
            pivotCacheRecord15.Append(missingItem173);
            pivotCacheRecord15.Append(missingItem174);
            pivotCacheRecord15.Append(missingItem175);
            pivotCacheRecord15.Append(missingItem176);
            pivotCacheRecord15.Append(missingItem177);
            pivotCacheRecord15.Append(missingItem178);
            pivotCacheRecord15.Append(missingItem179);
            pivotCacheRecord15.Append(missingItem180);
            pivotCacheRecord15.Append(missingItem181);
            pivotCacheRecord15.Append(missingItem182);
            pivotCacheRecord15.Append(missingItem183);
            pivotCacheRecord15.Append(fieldItem45);

            PivotCacheRecord pivotCacheRecord16 = new PivotCacheRecord();
            MissingItem missingItem184 = new MissingItem();
            MissingItem missingItem185 = new MissingItem();
            FieldItem fieldItem46 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem186 = new MissingItem();
            MissingItem missingItem187 = new MissingItem();
            FieldItem fieldItem47 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem188 = new MissingItem();
            MissingItem missingItem189 = new MissingItem();
            MissingItem missingItem190 = new MissingItem();
            MissingItem missingItem191 = new MissingItem();
            MissingItem missingItem192 = new MissingItem();
            MissingItem missingItem193 = new MissingItem();
            MissingItem missingItem194 = new MissingItem();
            MissingItem missingItem195 = new MissingItem();
            MissingItem missingItem196 = new MissingItem();
            MissingItem missingItem197 = new MissingItem();
            MissingItem missingItem198 = new MissingItem();
            MissingItem missingItem199 = new MissingItem();
            MissingItem missingItem200 = new MissingItem();
            MissingItem missingItem201 = new MissingItem();
            MissingItem missingItem202 = new MissingItem();
            MissingItem missingItem203 = new MissingItem();
            MissingItem missingItem204 = new MissingItem();
            MissingItem missingItem205 = new MissingItem();
            MissingItem missingItem206 = new MissingItem();
            MissingItem missingItem207 = new MissingItem();
            FieldItem fieldItem48 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord16.Append(missingItem184);
            pivotCacheRecord16.Append(missingItem185);
            pivotCacheRecord16.Append(fieldItem46);
            pivotCacheRecord16.Append(missingItem186);
            pivotCacheRecord16.Append(missingItem187);
            pivotCacheRecord16.Append(fieldItem47);
            pivotCacheRecord16.Append(missingItem188);
            pivotCacheRecord16.Append(missingItem189);
            pivotCacheRecord16.Append(missingItem190);
            pivotCacheRecord16.Append(missingItem191);
            pivotCacheRecord16.Append(missingItem192);
            pivotCacheRecord16.Append(missingItem193);
            pivotCacheRecord16.Append(missingItem194);
            pivotCacheRecord16.Append(missingItem195);
            pivotCacheRecord16.Append(missingItem196);
            pivotCacheRecord16.Append(missingItem197);
            pivotCacheRecord16.Append(missingItem198);
            pivotCacheRecord16.Append(missingItem199);
            pivotCacheRecord16.Append(missingItem200);
            pivotCacheRecord16.Append(missingItem201);
            pivotCacheRecord16.Append(missingItem202);
            pivotCacheRecord16.Append(missingItem203);
            pivotCacheRecord16.Append(missingItem204);
            pivotCacheRecord16.Append(missingItem205);
            pivotCacheRecord16.Append(missingItem206);
            pivotCacheRecord16.Append(missingItem207);
            pivotCacheRecord16.Append(fieldItem48);

            PivotCacheRecord pivotCacheRecord17 = new PivotCacheRecord();
            MissingItem missingItem208 = new MissingItem();
            MissingItem missingItem209 = new MissingItem();
            FieldItem fieldItem49 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem210 = new MissingItem();
            MissingItem missingItem211 = new MissingItem();
            FieldItem fieldItem50 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem212 = new MissingItem();
            MissingItem missingItem213 = new MissingItem();
            MissingItem missingItem214 = new MissingItem();
            MissingItem missingItem215 = new MissingItem();
            MissingItem missingItem216 = new MissingItem();
            MissingItem missingItem217 = new MissingItem();
            MissingItem missingItem218 = new MissingItem();
            MissingItem missingItem219 = new MissingItem();
            MissingItem missingItem220 = new MissingItem();
            MissingItem missingItem221 = new MissingItem();
            MissingItem missingItem222 = new MissingItem();
            MissingItem missingItem223 = new MissingItem();
            MissingItem missingItem224 = new MissingItem();
            MissingItem missingItem225 = new MissingItem();
            MissingItem missingItem226 = new MissingItem();
            MissingItem missingItem227 = new MissingItem();
            MissingItem missingItem228 = new MissingItem();
            MissingItem missingItem229 = new MissingItem();
            MissingItem missingItem230 = new MissingItem();
            MissingItem missingItem231 = new MissingItem();
            FieldItem fieldItem51 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord17.Append(missingItem208);
            pivotCacheRecord17.Append(missingItem209);
            pivotCacheRecord17.Append(fieldItem49);
            pivotCacheRecord17.Append(missingItem210);
            pivotCacheRecord17.Append(missingItem211);
            pivotCacheRecord17.Append(fieldItem50);
            pivotCacheRecord17.Append(missingItem212);
            pivotCacheRecord17.Append(missingItem213);
            pivotCacheRecord17.Append(missingItem214);
            pivotCacheRecord17.Append(missingItem215);
            pivotCacheRecord17.Append(missingItem216);
            pivotCacheRecord17.Append(missingItem217);
            pivotCacheRecord17.Append(missingItem218);
            pivotCacheRecord17.Append(missingItem219);
            pivotCacheRecord17.Append(missingItem220);
            pivotCacheRecord17.Append(missingItem221);
            pivotCacheRecord17.Append(missingItem222);
            pivotCacheRecord17.Append(missingItem223);
            pivotCacheRecord17.Append(missingItem224);
            pivotCacheRecord17.Append(missingItem225);
            pivotCacheRecord17.Append(missingItem226);
            pivotCacheRecord17.Append(missingItem227);
            pivotCacheRecord17.Append(missingItem228);
            pivotCacheRecord17.Append(missingItem229);
            pivotCacheRecord17.Append(missingItem230);
            pivotCacheRecord17.Append(missingItem231);
            pivotCacheRecord17.Append(fieldItem51);

            PivotCacheRecord pivotCacheRecord18 = new PivotCacheRecord();
            MissingItem missingItem232 = new MissingItem();
            MissingItem missingItem233 = new MissingItem();
            FieldItem fieldItem52 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem234 = new MissingItem();
            MissingItem missingItem235 = new MissingItem();
            FieldItem fieldItem53 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem236 = new MissingItem();
            MissingItem missingItem237 = new MissingItem();
            MissingItem missingItem238 = new MissingItem();
            MissingItem missingItem239 = new MissingItem();
            MissingItem missingItem240 = new MissingItem();
            MissingItem missingItem241 = new MissingItem();
            MissingItem missingItem242 = new MissingItem();
            MissingItem missingItem243 = new MissingItem();
            MissingItem missingItem244 = new MissingItem();
            MissingItem missingItem245 = new MissingItem();
            MissingItem missingItem246 = new MissingItem();
            MissingItem missingItem247 = new MissingItem();
            MissingItem missingItem248 = new MissingItem();
            MissingItem missingItem249 = new MissingItem();
            MissingItem missingItem250 = new MissingItem();
            MissingItem missingItem251 = new MissingItem();
            MissingItem missingItem252 = new MissingItem();
            MissingItem missingItem253 = new MissingItem();
            MissingItem missingItem254 = new MissingItem();
            MissingItem missingItem255 = new MissingItem();
            FieldItem fieldItem54 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord18.Append(missingItem232);
            pivotCacheRecord18.Append(missingItem233);
            pivotCacheRecord18.Append(fieldItem52);
            pivotCacheRecord18.Append(missingItem234);
            pivotCacheRecord18.Append(missingItem235);
            pivotCacheRecord18.Append(fieldItem53);
            pivotCacheRecord18.Append(missingItem236);
            pivotCacheRecord18.Append(missingItem237);
            pivotCacheRecord18.Append(missingItem238);
            pivotCacheRecord18.Append(missingItem239);
            pivotCacheRecord18.Append(missingItem240);
            pivotCacheRecord18.Append(missingItem241);
            pivotCacheRecord18.Append(missingItem242);
            pivotCacheRecord18.Append(missingItem243);
            pivotCacheRecord18.Append(missingItem244);
            pivotCacheRecord18.Append(missingItem245);
            pivotCacheRecord18.Append(missingItem246);
            pivotCacheRecord18.Append(missingItem247);
            pivotCacheRecord18.Append(missingItem248);
            pivotCacheRecord18.Append(missingItem249);
            pivotCacheRecord18.Append(missingItem250);
            pivotCacheRecord18.Append(missingItem251);
            pivotCacheRecord18.Append(missingItem252);
            pivotCacheRecord18.Append(missingItem253);
            pivotCacheRecord18.Append(missingItem254);
            pivotCacheRecord18.Append(missingItem255);
            pivotCacheRecord18.Append(fieldItem54);

            PivotCacheRecord pivotCacheRecord19 = new PivotCacheRecord();
            MissingItem missingItem256 = new MissingItem();
            MissingItem missingItem257 = new MissingItem();
            FieldItem fieldItem55 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem258 = new MissingItem();
            MissingItem missingItem259 = new MissingItem();
            FieldItem fieldItem56 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem260 = new MissingItem();
            MissingItem missingItem261 = new MissingItem();
            MissingItem missingItem262 = new MissingItem();
            MissingItem missingItem263 = new MissingItem();
            MissingItem missingItem264 = new MissingItem();
            MissingItem missingItem265 = new MissingItem();
            MissingItem missingItem266 = new MissingItem();
            MissingItem missingItem267 = new MissingItem();
            MissingItem missingItem268 = new MissingItem();
            MissingItem missingItem269 = new MissingItem();
            MissingItem missingItem270 = new MissingItem();
            MissingItem missingItem271 = new MissingItem();
            MissingItem missingItem272 = new MissingItem();
            MissingItem missingItem273 = new MissingItem();
            MissingItem missingItem274 = new MissingItem();
            MissingItem missingItem275 = new MissingItem();
            MissingItem missingItem276 = new MissingItem();
            MissingItem missingItem277 = new MissingItem();
            MissingItem missingItem278 = new MissingItem();
            MissingItem missingItem279 = new MissingItem();
            FieldItem fieldItem57 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord19.Append(missingItem256);
            pivotCacheRecord19.Append(missingItem257);
            pivotCacheRecord19.Append(fieldItem55);
            pivotCacheRecord19.Append(missingItem258);
            pivotCacheRecord19.Append(missingItem259);
            pivotCacheRecord19.Append(fieldItem56);
            pivotCacheRecord19.Append(missingItem260);
            pivotCacheRecord19.Append(missingItem261);
            pivotCacheRecord19.Append(missingItem262);
            pivotCacheRecord19.Append(missingItem263);
            pivotCacheRecord19.Append(missingItem264);
            pivotCacheRecord19.Append(missingItem265);
            pivotCacheRecord19.Append(missingItem266);
            pivotCacheRecord19.Append(missingItem267);
            pivotCacheRecord19.Append(missingItem268);
            pivotCacheRecord19.Append(missingItem269);
            pivotCacheRecord19.Append(missingItem270);
            pivotCacheRecord19.Append(missingItem271);
            pivotCacheRecord19.Append(missingItem272);
            pivotCacheRecord19.Append(missingItem273);
            pivotCacheRecord19.Append(missingItem274);
            pivotCacheRecord19.Append(missingItem275);
            pivotCacheRecord19.Append(missingItem276);
            pivotCacheRecord19.Append(missingItem277);
            pivotCacheRecord19.Append(missingItem278);
            pivotCacheRecord19.Append(missingItem279);
            pivotCacheRecord19.Append(fieldItem57);

            PivotCacheRecord pivotCacheRecord20 = new PivotCacheRecord();
            MissingItem missingItem280 = new MissingItem();
            MissingItem missingItem281 = new MissingItem();
            FieldItem fieldItem58 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem282 = new MissingItem();
            MissingItem missingItem283 = new MissingItem();
            FieldItem fieldItem59 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem284 = new MissingItem();
            MissingItem missingItem285 = new MissingItem();
            MissingItem missingItem286 = new MissingItem();
            MissingItem missingItem287 = new MissingItem();
            MissingItem missingItem288 = new MissingItem();
            MissingItem missingItem289 = new MissingItem();
            MissingItem missingItem290 = new MissingItem();
            MissingItem missingItem291 = new MissingItem();
            MissingItem missingItem292 = new MissingItem();
            MissingItem missingItem293 = new MissingItem();
            MissingItem missingItem294 = new MissingItem();
            MissingItem missingItem295 = new MissingItem();
            MissingItem missingItem296 = new MissingItem();
            MissingItem missingItem297 = new MissingItem();
            MissingItem missingItem298 = new MissingItem();
            MissingItem missingItem299 = new MissingItem();
            MissingItem missingItem300 = new MissingItem();
            MissingItem missingItem301 = new MissingItem();
            MissingItem missingItem302 = new MissingItem();
            MissingItem missingItem303 = new MissingItem();
            FieldItem fieldItem60 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord20.Append(missingItem280);
            pivotCacheRecord20.Append(missingItem281);
            pivotCacheRecord20.Append(fieldItem58);
            pivotCacheRecord20.Append(missingItem282);
            pivotCacheRecord20.Append(missingItem283);
            pivotCacheRecord20.Append(fieldItem59);
            pivotCacheRecord20.Append(missingItem284);
            pivotCacheRecord20.Append(missingItem285);
            pivotCacheRecord20.Append(missingItem286);
            pivotCacheRecord20.Append(missingItem287);
            pivotCacheRecord20.Append(missingItem288);
            pivotCacheRecord20.Append(missingItem289);
            pivotCacheRecord20.Append(missingItem290);
            pivotCacheRecord20.Append(missingItem291);
            pivotCacheRecord20.Append(missingItem292);
            pivotCacheRecord20.Append(missingItem293);
            pivotCacheRecord20.Append(missingItem294);
            pivotCacheRecord20.Append(missingItem295);
            pivotCacheRecord20.Append(missingItem296);
            pivotCacheRecord20.Append(missingItem297);
            pivotCacheRecord20.Append(missingItem298);
            pivotCacheRecord20.Append(missingItem299);
            pivotCacheRecord20.Append(missingItem300);
            pivotCacheRecord20.Append(missingItem301);
            pivotCacheRecord20.Append(missingItem302);
            pivotCacheRecord20.Append(missingItem303);
            pivotCacheRecord20.Append(fieldItem60);

            PivotCacheRecord pivotCacheRecord21 = new PivotCacheRecord();
            MissingItem missingItem304 = new MissingItem();
            MissingItem missingItem305 = new MissingItem();
            FieldItem fieldItem61 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem306 = new MissingItem();
            MissingItem missingItem307 = new MissingItem();
            FieldItem fieldItem62 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem308 = new MissingItem();
            MissingItem missingItem309 = new MissingItem();
            MissingItem missingItem310 = new MissingItem();
            MissingItem missingItem311 = new MissingItem();
            MissingItem missingItem312 = new MissingItem();
            MissingItem missingItem313 = new MissingItem();
            MissingItem missingItem314 = new MissingItem();
            MissingItem missingItem315 = new MissingItem();
            MissingItem missingItem316 = new MissingItem();
            MissingItem missingItem317 = new MissingItem();
            MissingItem missingItem318 = new MissingItem();
            MissingItem missingItem319 = new MissingItem();
            MissingItem missingItem320 = new MissingItem();
            MissingItem missingItem321 = new MissingItem();
            MissingItem missingItem322 = new MissingItem();
            MissingItem missingItem323 = new MissingItem();
            MissingItem missingItem324 = new MissingItem();
            MissingItem missingItem325 = new MissingItem();
            MissingItem missingItem326 = new MissingItem();
            MissingItem missingItem327 = new MissingItem();
            FieldItem fieldItem63 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord21.Append(missingItem304);
            pivotCacheRecord21.Append(missingItem305);
            pivotCacheRecord21.Append(fieldItem61);
            pivotCacheRecord21.Append(missingItem306);
            pivotCacheRecord21.Append(missingItem307);
            pivotCacheRecord21.Append(fieldItem62);
            pivotCacheRecord21.Append(missingItem308);
            pivotCacheRecord21.Append(missingItem309);
            pivotCacheRecord21.Append(missingItem310);
            pivotCacheRecord21.Append(missingItem311);
            pivotCacheRecord21.Append(missingItem312);
            pivotCacheRecord21.Append(missingItem313);
            pivotCacheRecord21.Append(missingItem314);
            pivotCacheRecord21.Append(missingItem315);
            pivotCacheRecord21.Append(missingItem316);
            pivotCacheRecord21.Append(missingItem317);
            pivotCacheRecord21.Append(missingItem318);
            pivotCacheRecord21.Append(missingItem319);
            pivotCacheRecord21.Append(missingItem320);
            pivotCacheRecord21.Append(missingItem321);
            pivotCacheRecord21.Append(missingItem322);
            pivotCacheRecord21.Append(missingItem323);
            pivotCacheRecord21.Append(missingItem324);
            pivotCacheRecord21.Append(missingItem325);
            pivotCacheRecord21.Append(missingItem326);
            pivotCacheRecord21.Append(missingItem327);
            pivotCacheRecord21.Append(fieldItem63);

            PivotCacheRecord pivotCacheRecord22 = new PivotCacheRecord();
            MissingItem missingItem328 = new MissingItem();
            MissingItem missingItem329 = new MissingItem();
            FieldItem fieldItem64 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem330 = new MissingItem();
            MissingItem missingItem331 = new MissingItem();
            FieldItem fieldItem65 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem332 = new MissingItem();
            MissingItem missingItem333 = new MissingItem();
            MissingItem missingItem334 = new MissingItem();
            MissingItem missingItem335 = new MissingItem();
            MissingItem missingItem336 = new MissingItem();
            MissingItem missingItem337 = new MissingItem();
            MissingItem missingItem338 = new MissingItem();
            MissingItem missingItem339 = new MissingItem();
            MissingItem missingItem340 = new MissingItem();
            MissingItem missingItem341 = new MissingItem();
            MissingItem missingItem342 = new MissingItem();
            MissingItem missingItem343 = new MissingItem();
            MissingItem missingItem344 = new MissingItem();
            MissingItem missingItem345 = new MissingItem();
            MissingItem missingItem346 = new MissingItem();
            MissingItem missingItem347 = new MissingItem();
            MissingItem missingItem348 = new MissingItem();
            MissingItem missingItem349 = new MissingItem();
            MissingItem missingItem350 = new MissingItem();
            MissingItem missingItem351 = new MissingItem();
            FieldItem fieldItem66 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord22.Append(missingItem328);
            pivotCacheRecord22.Append(missingItem329);
            pivotCacheRecord22.Append(fieldItem64);
            pivotCacheRecord22.Append(missingItem330);
            pivotCacheRecord22.Append(missingItem331);
            pivotCacheRecord22.Append(fieldItem65);
            pivotCacheRecord22.Append(missingItem332);
            pivotCacheRecord22.Append(missingItem333);
            pivotCacheRecord22.Append(missingItem334);
            pivotCacheRecord22.Append(missingItem335);
            pivotCacheRecord22.Append(missingItem336);
            pivotCacheRecord22.Append(missingItem337);
            pivotCacheRecord22.Append(missingItem338);
            pivotCacheRecord22.Append(missingItem339);
            pivotCacheRecord22.Append(missingItem340);
            pivotCacheRecord22.Append(missingItem341);
            pivotCacheRecord22.Append(missingItem342);
            pivotCacheRecord22.Append(missingItem343);
            pivotCacheRecord22.Append(missingItem344);
            pivotCacheRecord22.Append(missingItem345);
            pivotCacheRecord22.Append(missingItem346);
            pivotCacheRecord22.Append(missingItem347);
            pivotCacheRecord22.Append(missingItem348);
            pivotCacheRecord22.Append(missingItem349);
            pivotCacheRecord22.Append(missingItem350);
            pivotCacheRecord22.Append(missingItem351);
            pivotCacheRecord22.Append(fieldItem66);

            PivotCacheRecord pivotCacheRecord23 = new PivotCacheRecord();
            MissingItem missingItem352 = new MissingItem();
            MissingItem missingItem353 = new MissingItem();
            FieldItem fieldItem67 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem354 = new MissingItem();
            MissingItem missingItem355 = new MissingItem();
            FieldItem fieldItem68 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem356 = new MissingItem();
            MissingItem missingItem357 = new MissingItem();
            MissingItem missingItem358 = new MissingItem();
            MissingItem missingItem359 = new MissingItem();
            MissingItem missingItem360 = new MissingItem();
            MissingItem missingItem361 = new MissingItem();
            MissingItem missingItem362 = new MissingItem();
            MissingItem missingItem363 = new MissingItem();
            MissingItem missingItem364 = new MissingItem();
            MissingItem missingItem365 = new MissingItem();
            MissingItem missingItem366 = new MissingItem();
            MissingItem missingItem367 = new MissingItem();
            MissingItem missingItem368 = new MissingItem();
            MissingItem missingItem369 = new MissingItem();
            MissingItem missingItem370 = new MissingItem();
            MissingItem missingItem371 = new MissingItem();
            MissingItem missingItem372 = new MissingItem();
            MissingItem missingItem373 = new MissingItem();
            MissingItem missingItem374 = new MissingItem();
            MissingItem missingItem375 = new MissingItem();
            FieldItem fieldItem69 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord23.Append(missingItem352);
            pivotCacheRecord23.Append(missingItem353);
            pivotCacheRecord23.Append(fieldItem67);
            pivotCacheRecord23.Append(missingItem354);
            pivotCacheRecord23.Append(missingItem355);
            pivotCacheRecord23.Append(fieldItem68);
            pivotCacheRecord23.Append(missingItem356);
            pivotCacheRecord23.Append(missingItem357);
            pivotCacheRecord23.Append(missingItem358);
            pivotCacheRecord23.Append(missingItem359);
            pivotCacheRecord23.Append(missingItem360);
            pivotCacheRecord23.Append(missingItem361);
            pivotCacheRecord23.Append(missingItem362);
            pivotCacheRecord23.Append(missingItem363);
            pivotCacheRecord23.Append(missingItem364);
            pivotCacheRecord23.Append(missingItem365);
            pivotCacheRecord23.Append(missingItem366);
            pivotCacheRecord23.Append(missingItem367);
            pivotCacheRecord23.Append(missingItem368);
            pivotCacheRecord23.Append(missingItem369);
            pivotCacheRecord23.Append(missingItem370);
            pivotCacheRecord23.Append(missingItem371);
            pivotCacheRecord23.Append(missingItem372);
            pivotCacheRecord23.Append(missingItem373);
            pivotCacheRecord23.Append(missingItem374);
            pivotCacheRecord23.Append(missingItem375);
            pivotCacheRecord23.Append(fieldItem69);

            PivotCacheRecord pivotCacheRecord24 = new PivotCacheRecord();
            MissingItem missingItem376 = new MissingItem();
            MissingItem missingItem377 = new MissingItem();
            FieldItem fieldItem70 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem378 = new MissingItem();
            MissingItem missingItem379 = new MissingItem();
            FieldItem fieldItem71 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem380 = new MissingItem();
            MissingItem missingItem381 = new MissingItem();
            MissingItem missingItem382 = new MissingItem();
            MissingItem missingItem383 = new MissingItem();
            MissingItem missingItem384 = new MissingItem();
            MissingItem missingItem385 = new MissingItem();
            MissingItem missingItem386 = new MissingItem();
            MissingItem missingItem387 = new MissingItem();
            MissingItem missingItem388 = new MissingItem();
            MissingItem missingItem389 = new MissingItem();
            MissingItem missingItem390 = new MissingItem();
            MissingItem missingItem391 = new MissingItem();
            MissingItem missingItem392 = new MissingItem();
            MissingItem missingItem393 = new MissingItem();
            MissingItem missingItem394 = new MissingItem();
            MissingItem missingItem395 = new MissingItem();
            MissingItem missingItem396 = new MissingItem();
            MissingItem missingItem397 = new MissingItem();
            MissingItem missingItem398 = new MissingItem();
            MissingItem missingItem399 = new MissingItem();
            FieldItem fieldItem72 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord24.Append(missingItem376);
            pivotCacheRecord24.Append(missingItem377);
            pivotCacheRecord24.Append(fieldItem70);
            pivotCacheRecord24.Append(missingItem378);
            pivotCacheRecord24.Append(missingItem379);
            pivotCacheRecord24.Append(fieldItem71);
            pivotCacheRecord24.Append(missingItem380);
            pivotCacheRecord24.Append(missingItem381);
            pivotCacheRecord24.Append(missingItem382);
            pivotCacheRecord24.Append(missingItem383);
            pivotCacheRecord24.Append(missingItem384);
            pivotCacheRecord24.Append(missingItem385);
            pivotCacheRecord24.Append(missingItem386);
            pivotCacheRecord24.Append(missingItem387);
            pivotCacheRecord24.Append(missingItem388);
            pivotCacheRecord24.Append(missingItem389);
            pivotCacheRecord24.Append(missingItem390);
            pivotCacheRecord24.Append(missingItem391);
            pivotCacheRecord24.Append(missingItem392);
            pivotCacheRecord24.Append(missingItem393);
            pivotCacheRecord24.Append(missingItem394);
            pivotCacheRecord24.Append(missingItem395);
            pivotCacheRecord24.Append(missingItem396);
            pivotCacheRecord24.Append(missingItem397);
            pivotCacheRecord24.Append(missingItem398);
            pivotCacheRecord24.Append(missingItem399);
            pivotCacheRecord24.Append(fieldItem72);

            PivotCacheRecord pivotCacheRecord25 = new PivotCacheRecord();
            MissingItem missingItem400 = new MissingItem();
            MissingItem missingItem401 = new MissingItem();
            FieldItem fieldItem73 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem402 = new MissingItem();
            MissingItem missingItem403 = new MissingItem();
            FieldItem fieldItem74 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem404 = new MissingItem();
            MissingItem missingItem405 = new MissingItem();
            MissingItem missingItem406 = new MissingItem();
            MissingItem missingItem407 = new MissingItem();
            MissingItem missingItem408 = new MissingItem();
            MissingItem missingItem409 = new MissingItem();
            MissingItem missingItem410 = new MissingItem();
            MissingItem missingItem411 = new MissingItem();
            MissingItem missingItem412 = new MissingItem();
            MissingItem missingItem413 = new MissingItem();
            MissingItem missingItem414 = new MissingItem();
            MissingItem missingItem415 = new MissingItem();
            MissingItem missingItem416 = new MissingItem();
            MissingItem missingItem417 = new MissingItem();
            MissingItem missingItem418 = new MissingItem();
            MissingItem missingItem419 = new MissingItem();
            MissingItem missingItem420 = new MissingItem();
            MissingItem missingItem421 = new MissingItem();
            MissingItem missingItem422 = new MissingItem();
            MissingItem missingItem423 = new MissingItem();
            FieldItem fieldItem75 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord25.Append(missingItem400);
            pivotCacheRecord25.Append(missingItem401);
            pivotCacheRecord25.Append(fieldItem73);
            pivotCacheRecord25.Append(missingItem402);
            pivotCacheRecord25.Append(missingItem403);
            pivotCacheRecord25.Append(fieldItem74);
            pivotCacheRecord25.Append(missingItem404);
            pivotCacheRecord25.Append(missingItem405);
            pivotCacheRecord25.Append(missingItem406);
            pivotCacheRecord25.Append(missingItem407);
            pivotCacheRecord25.Append(missingItem408);
            pivotCacheRecord25.Append(missingItem409);
            pivotCacheRecord25.Append(missingItem410);
            pivotCacheRecord25.Append(missingItem411);
            pivotCacheRecord25.Append(missingItem412);
            pivotCacheRecord25.Append(missingItem413);
            pivotCacheRecord25.Append(missingItem414);
            pivotCacheRecord25.Append(missingItem415);
            pivotCacheRecord25.Append(missingItem416);
            pivotCacheRecord25.Append(missingItem417);
            pivotCacheRecord25.Append(missingItem418);
            pivotCacheRecord25.Append(missingItem419);
            pivotCacheRecord25.Append(missingItem420);
            pivotCacheRecord25.Append(missingItem421);
            pivotCacheRecord25.Append(missingItem422);
            pivotCacheRecord25.Append(missingItem423);
            pivotCacheRecord25.Append(fieldItem75);

            PivotCacheRecord pivotCacheRecord26 = new PivotCacheRecord();
            MissingItem missingItem424 = new MissingItem();
            MissingItem missingItem425 = new MissingItem();
            FieldItem fieldItem76 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem426 = new MissingItem();
            MissingItem missingItem427 = new MissingItem();
            FieldItem fieldItem77 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem428 = new MissingItem();
            MissingItem missingItem429 = new MissingItem();
            MissingItem missingItem430 = new MissingItem();
            MissingItem missingItem431 = new MissingItem();
            MissingItem missingItem432 = new MissingItem();
            MissingItem missingItem433 = new MissingItem();
            MissingItem missingItem434 = new MissingItem();
            MissingItem missingItem435 = new MissingItem();
            MissingItem missingItem436 = new MissingItem();
            MissingItem missingItem437 = new MissingItem();
            MissingItem missingItem438 = new MissingItem();
            MissingItem missingItem439 = new MissingItem();
            MissingItem missingItem440 = new MissingItem();
            MissingItem missingItem441 = new MissingItem();
            MissingItem missingItem442 = new MissingItem();
            MissingItem missingItem443 = new MissingItem();
            MissingItem missingItem444 = new MissingItem();
            MissingItem missingItem445 = new MissingItem();
            MissingItem missingItem446 = new MissingItem();
            MissingItem missingItem447 = new MissingItem();
            FieldItem fieldItem78 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord26.Append(missingItem424);
            pivotCacheRecord26.Append(missingItem425);
            pivotCacheRecord26.Append(fieldItem76);
            pivotCacheRecord26.Append(missingItem426);
            pivotCacheRecord26.Append(missingItem427);
            pivotCacheRecord26.Append(fieldItem77);
            pivotCacheRecord26.Append(missingItem428);
            pivotCacheRecord26.Append(missingItem429);
            pivotCacheRecord26.Append(missingItem430);
            pivotCacheRecord26.Append(missingItem431);
            pivotCacheRecord26.Append(missingItem432);
            pivotCacheRecord26.Append(missingItem433);
            pivotCacheRecord26.Append(missingItem434);
            pivotCacheRecord26.Append(missingItem435);
            pivotCacheRecord26.Append(missingItem436);
            pivotCacheRecord26.Append(missingItem437);
            pivotCacheRecord26.Append(missingItem438);
            pivotCacheRecord26.Append(missingItem439);
            pivotCacheRecord26.Append(missingItem440);
            pivotCacheRecord26.Append(missingItem441);
            pivotCacheRecord26.Append(missingItem442);
            pivotCacheRecord26.Append(missingItem443);
            pivotCacheRecord26.Append(missingItem444);
            pivotCacheRecord26.Append(missingItem445);
            pivotCacheRecord26.Append(missingItem446);
            pivotCacheRecord26.Append(missingItem447);
            pivotCacheRecord26.Append(fieldItem78);

            PivotCacheRecord pivotCacheRecord27 = new PivotCacheRecord();
            MissingItem missingItem448 = new MissingItem();
            MissingItem missingItem449 = new MissingItem();
            FieldItem fieldItem79 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem450 = new MissingItem();
            MissingItem missingItem451 = new MissingItem();
            FieldItem fieldItem80 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem452 = new MissingItem();
            MissingItem missingItem453 = new MissingItem();
            MissingItem missingItem454 = new MissingItem();
            MissingItem missingItem455 = new MissingItem();
            MissingItem missingItem456 = new MissingItem();
            MissingItem missingItem457 = new MissingItem();
            MissingItem missingItem458 = new MissingItem();
            MissingItem missingItem459 = new MissingItem();
            MissingItem missingItem460 = new MissingItem();
            MissingItem missingItem461 = new MissingItem();
            MissingItem missingItem462 = new MissingItem();
            MissingItem missingItem463 = new MissingItem();
            MissingItem missingItem464 = new MissingItem();
            MissingItem missingItem465 = new MissingItem();
            MissingItem missingItem466 = new MissingItem();
            MissingItem missingItem467 = new MissingItem();
            MissingItem missingItem468 = new MissingItem();
            MissingItem missingItem469 = new MissingItem();
            MissingItem missingItem470 = new MissingItem();
            MissingItem missingItem471 = new MissingItem();
            FieldItem fieldItem81 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord27.Append(missingItem448);
            pivotCacheRecord27.Append(missingItem449);
            pivotCacheRecord27.Append(fieldItem79);
            pivotCacheRecord27.Append(missingItem450);
            pivotCacheRecord27.Append(missingItem451);
            pivotCacheRecord27.Append(fieldItem80);
            pivotCacheRecord27.Append(missingItem452);
            pivotCacheRecord27.Append(missingItem453);
            pivotCacheRecord27.Append(missingItem454);
            pivotCacheRecord27.Append(missingItem455);
            pivotCacheRecord27.Append(missingItem456);
            pivotCacheRecord27.Append(missingItem457);
            pivotCacheRecord27.Append(missingItem458);
            pivotCacheRecord27.Append(missingItem459);
            pivotCacheRecord27.Append(missingItem460);
            pivotCacheRecord27.Append(missingItem461);
            pivotCacheRecord27.Append(missingItem462);
            pivotCacheRecord27.Append(missingItem463);
            pivotCacheRecord27.Append(missingItem464);
            pivotCacheRecord27.Append(missingItem465);
            pivotCacheRecord27.Append(missingItem466);
            pivotCacheRecord27.Append(missingItem467);
            pivotCacheRecord27.Append(missingItem468);
            pivotCacheRecord27.Append(missingItem469);
            pivotCacheRecord27.Append(missingItem470);
            pivotCacheRecord27.Append(missingItem471);
            pivotCacheRecord27.Append(fieldItem81);

            PivotCacheRecord pivotCacheRecord28 = new PivotCacheRecord();
            MissingItem missingItem472 = new MissingItem();
            MissingItem missingItem473 = new MissingItem();
            FieldItem fieldItem82 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem474 = new MissingItem();
            MissingItem missingItem475 = new MissingItem();
            FieldItem fieldItem83 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem476 = new MissingItem();
            MissingItem missingItem477 = new MissingItem();
            MissingItem missingItem478 = new MissingItem();
            MissingItem missingItem479 = new MissingItem();
            MissingItem missingItem480 = new MissingItem();
            MissingItem missingItem481 = new MissingItem();
            MissingItem missingItem482 = new MissingItem();
            MissingItem missingItem483 = new MissingItem();
            MissingItem missingItem484 = new MissingItem();
            MissingItem missingItem485 = new MissingItem();
            MissingItem missingItem486 = new MissingItem();
            MissingItem missingItem487 = new MissingItem();
            MissingItem missingItem488 = new MissingItem();
            MissingItem missingItem489 = new MissingItem();
            MissingItem missingItem490 = new MissingItem();
            MissingItem missingItem491 = new MissingItem();
            MissingItem missingItem492 = new MissingItem();
            MissingItem missingItem493 = new MissingItem();
            MissingItem missingItem494 = new MissingItem();
            MissingItem missingItem495 = new MissingItem();
            FieldItem fieldItem84 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord28.Append(missingItem472);
            pivotCacheRecord28.Append(missingItem473);
            pivotCacheRecord28.Append(fieldItem82);
            pivotCacheRecord28.Append(missingItem474);
            pivotCacheRecord28.Append(missingItem475);
            pivotCacheRecord28.Append(fieldItem83);
            pivotCacheRecord28.Append(missingItem476);
            pivotCacheRecord28.Append(missingItem477);
            pivotCacheRecord28.Append(missingItem478);
            pivotCacheRecord28.Append(missingItem479);
            pivotCacheRecord28.Append(missingItem480);
            pivotCacheRecord28.Append(missingItem481);
            pivotCacheRecord28.Append(missingItem482);
            pivotCacheRecord28.Append(missingItem483);
            pivotCacheRecord28.Append(missingItem484);
            pivotCacheRecord28.Append(missingItem485);
            pivotCacheRecord28.Append(missingItem486);
            pivotCacheRecord28.Append(missingItem487);
            pivotCacheRecord28.Append(missingItem488);
            pivotCacheRecord28.Append(missingItem489);
            pivotCacheRecord28.Append(missingItem490);
            pivotCacheRecord28.Append(missingItem491);
            pivotCacheRecord28.Append(missingItem492);
            pivotCacheRecord28.Append(missingItem493);
            pivotCacheRecord28.Append(missingItem494);
            pivotCacheRecord28.Append(missingItem495);
            pivotCacheRecord28.Append(fieldItem84);

            PivotCacheRecord pivotCacheRecord29 = new PivotCacheRecord();
            MissingItem missingItem496 = new MissingItem();
            MissingItem missingItem497 = new MissingItem();
            FieldItem fieldItem85 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem498 = new MissingItem();
            MissingItem missingItem499 = new MissingItem();
            FieldItem fieldItem86 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem500 = new MissingItem();
            MissingItem missingItem501 = new MissingItem();
            MissingItem missingItem502 = new MissingItem();
            MissingItem missingItem503 = new MissingItem();
            MissingItem missingItem504 = new MissingItem();
            MissingItem missingItem505 = new MissingItem();
            MissingItem missingItem506 = new MissingItem();
            MissingItem missingItem507 = new MissingItem();
            MissingItem missingItem508 = new MissingItem();
            MissingItem missingItem509 = new MissingItem();
            MissingItem missingItem510 = new MissingItem();
            MissingItem missingItem511 = new MissingItem();
            MissingItem missingItem512 = new MissingItem();
            MissingItem missingItem513 = new MissingItem();
            MissingItem missingItem514 = new MissingItem();
            MissingItem missingItem515 = new MissingItem();
            MissingItem missingItem516 = new MissingItem();
            MissingItem missingItem517 = new MissingItem();
            MissingItem missingItem518 = new MissingItem();
            MissingItem missingItem519 = new MissingItem();
            FieldItem fieldItem87 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord29.Append(missingItem496);
            pivotCacheRecord29.Append(missingItem497);
            pivotCacheRecord29.Append(fieldItem85);
            pivotCacheRecord29.Append(missingItem498);
            pivotCacheRecord29.Append(missingItem499);
            pivotCacheRecord29.Append(fieldItem86);
            pivotCacheRecord29.Append(missingItem500);
            pivotCacheRecord29.Append(missingItem501);
            pivotCacheRecord29.Append(missingItem502);
            pivotCacheRecord29.Append(missingItem503);
            pivotCacheRecord29.Append(missingItem504);
            pivotCacheRecord29.Append(missingItem505);
            pivotCacheRecord29.Append(missingItem506);
            pivotCacheRecord29.Append(missingItem507);
            pivotCacheRecord29.Append(missingItem508);
            pivotCacheRecord29.Append(missingItem509);
            pivotCacheRecord29.Append(missingItem510);
            pivotCacheRecord29.Append(missingItem511);
            pivotCacheRecord29.Append(missingItem512);
            pivotCacheRecord29.Append(missingItem513);
            pivotCacheRecord29.Append(missingItem514);
            pivotCacheRecord29.Append(missingItem515);
            pivotCacheRecord29.Append(missingItem516);
            pivotCacheRecord29.Append(missingItem517);
            pivotCacheRecord29.Append(missingItem518);
            pivotCacheRecord29.Append(missingItem519);
            pivotCacheRecord29.Append(fieldItem87);

            PivotCacheRecord pivotCacheRecord30 = new PivotCacheRecord();
            MissingItem missingItem520 = new MissingItem();
            MissingItem missingItem521 = new MissingItem();
            FieldItem fieldItem88 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem522 = new MissingItem();
            MissingItem missingItem523 = new MissingItem();
            FieldItem fieldItem89 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem524 = new MissingItem();
            MissingItem missingItem525 = new MissingItem();
            MissingItem missingItem526 = new MissingItem();
            MissingItem missingItem527 = new MissingItem();
            MissingItem missingItem528 = new MissingItem();
            MissingItem missingItem529 = new MissingItem();
            MissingItem missingItem530 = new MissingItem();
            MissingItem missingItem531 = new MissingItem();
            MissingItem missingItem532 = new MissingItem();
            MissingItem missingItem533 = new MissingItem();
            MissingItem missingItem534 = new MissingItem();
            MissingItem missingItem535 = new MissingItem();
            MissingItem missingItem536 = new MissingItem();
            MissingItem missingItem537 = new MissingItem();
            MissingItem missingItem538 = new MissingItem();
            MissingItem missingItem539 = new MissingItem();
            MissingItem missingItem540 = new MissingItem();
            MissingItem missingItem541 = new MissingItem();
            MissingItem missingItem542 = new MissingItem();
            MissingItem missingItem543 = new MissingItem();
            FieldItem fieldItem90 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord30.Append(missingItem520);
            pivotCacheRecord30.Append(missingItem521);
            pivotCacheRecord30.Append(fieldItem88);
            pivotCacheRecord30.Append(missingItem522);
            pivotCacheRecord30.Append(missingItem523);
            pivotCacheRecord30.Append(fieldItem89);
            pivotCacheRecord30.Append(missingItem524);
            pivotCacheRecord30.Append(missingItem525);
            pivotCacheRecord30.Append(missingItem526);
            pivotCacheRecord30.Append(missingItem527);
            pivotCacheRecord30.Append(missingItem528);
            pivotCacheRecord30.Append(missingItem529);
            pivotCacheRecord30.Append(missingItem530);
            pivotCacheRecord30.Append(missingItem531);
            pivotCacheRecord30.Append(missingItem532);
            pivotCacheRecord30.Append(missingItem533);
            pivotCacheRecord30.Append(missingItem534);
            pivotCacheRecord30.Append(missingItem535);
            pivotCacheRecord30.Append(missingItem536);
            pivotCacheRecord30.Append(missingItem537);
            pivotCacheRecord30.Append(missingItem538);
            pivotCacheRecord30.Append(missingItem539);
            pivotCacheRecord30.Append(missingItem540);
            pivotCacheRecord30.Append(missingItem541);
            pivotCacheRecord30.Append(missingItem542);
            pivotCacheRecord30.Append(missingItem543);
            pivotCacheRecord30.Append(fieldItem90);

            PivotCacheRecord pivotCacheRecord31 = new PivotCacheRecord();
            MissingItem missingItem544 = new MissingItem();
            MissingItem missingItem545 = new MissingItem();
            FieldItem fieldItem91 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem546 = new MissingItem();
            MissingItem missingItem547 = new MissingItem();
            FieldItem fieldItem92 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem548 = new MissingItem();
            MissingItem missingItem549 = new MissingItem();
            MissingItem missingItem550 = new MissingItem();
            MissingItem missingItem551 = new MissingItem();
            MissingItem missingItem552 = new MissingItem();
            MissingItem missingItem553 = new MissingItem();
            MissingItem missingItem554 = new MissingItem();
            MissingItem missingItem555 = new MissingItem();
            MissingItem missingItem556 = new MissingItem();
            MissingItem missingItem557 = new MissingItem();
            MissingItem missingItem558 = new MissingItem();
            MissingItem missingItem559 = new MissingItem();
            MissingItem missingItem560 = new MissingItem();
            MissingItem missingItem561 = new MissingItem();
            MissingItem missingItem562 = new MissingItem();
            MissingItem missingItem563 = new MissingItem();
            MissingItem missingItem564 = new MissingItem();
            MissingItem missingItem565 = new MissingItem();
            MissingItem missingItem566 = new MissingItem();
            MissingItem missingItem567 = new MissingItem();
            FieldItem fieldItem93 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord31.Append(missingItem544);
            pivotCacheRecord31.Append(missingItem545);
            pivotCacheRecord31.Append(fieldItem91);
            pivotCacheRecord31.Append(missingItem546);
            pivotCacheRecord31.Append(missingItem547);
            pivotCacheRecord31.Append(fieldItem92);
            pivotCacheRecord31.Append(missingItem548);
            pivotCacheRecord31.Append(missingItem549);
            pivotCacheRecord31.Append(missingItem550);
            pivotCacheRecord31.Append(missingItem551);
            pivotCacheRecord31.Append(missingItem552);
            pivotCacheRecord31.Append(missingItem553);
            pivotCacheRecord31.Append(missingItem554);
            pivotCacheRecord31.Append(missingItem555);
            pivotCacheRecord31.Append(missingItem556);
            pivotCacheRecord31.Append(missingItem557);
            pivotCacheRecord31.Append(missingItem558);
            pivotCacheRecord31.Append(missingItem559);
            pivotCacheRecord31.Append(missingItem560);
            pivotCacheRecord31.Append(missingItem561);
            pivotCacheRecord31.Append(missingItem562);
            pivotCacheRecord31.Append(missingItem563);
            pivotCacheRecord31.Append(missingItem564);
            pivotCacheRecord31.Append(missingItem565);
            pivotCacheRecord31.Append(missingItem566);
            pivotCacheRecord31.Append(missingItem567);
            pivotCacheRecord31.Append(fieldItem93);

            PivotCacheRecord pivotCacheRecord32 = new PivotCacheRecord();
            MissingItem missingItem568 = new MissingItem();
            MissingItem missingItem569 = new MissingItem();
            FieldItem fieldItem94 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem570 = new MissingItem();
            MissingItem missingItem571 = new MissingItem();
            FieldItem fieldItem95 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem572 = new MissingItem();
            MissingItem missingItem573 = new MissingItem();
            MissingItem missingItem574 = new MissingItem();
            MissingItem missingItem575 = new MissingItem();
            MissingItem missingItem576 = new MissingItem();
            MissingItem missingItem577 = new MissingItem();
            MissingItem missingItem578 = new MissingItem();
            MissingItem missingItem579 = new MissingItem();
            MissingItem missingItem580 = new MissingItem();
            MissingItem missingItem581 = new MissingItem();
            MissingItem missingItem582 = new MissingItem();
            MissingItem missingItem583 = new MissingItem();
            MissingItem missingItem584 = new MissingItem();
            MissingItem missingItem585 = new MissingItem();
            MissingItem missingItem586 = new MissingItem();
            MissingItem missingItem587 = new MissingItem();
            MissingItem missingItem588 = new MissingItem();
            MissingItem missingItem589 = new MissingItem();
            MissingItem missingItem590 = new MissingItem();
            MissingItem missingItem591 = new MissingItem();
            FieldItem fieldItem96 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord32.Append(missingItem568);
            pivotCacheRecord32.Append(missingItem569);
            pivotCacheRecord32.Append(fieldItem94);
            pivotCacheRecord32.Append(missingItem570);
            pivotCacheRecord32.Append(missingItem571);
            pivotCacheRecord32.Append(fieldItem95);
            pivotCacheRecord32.Append(missingItem572);
            pivotCacheRecord32.Append(missingItem573);
            pivotCacheRecord32.Append(missingItem574);
            pivotCacheRecord32.Append(missingItem575);
            pivotCacheRecord32.Append(missingItem576);
            pivotCacheRecord32.Append(missingItem577);
            pivotCacheRecord32.Append(missingItem578);
            pivotCacheRecord32.Append(missingItem579);
            pivotCacheRecord32.Append(missingItem580);
            pivotCacheRecord32.Append(missingItem581);
            pivotCacheRecord32.Append(missingItem582);
            pivotCacheRecord32.Append(missingItem583);
            pivotCacheRecord32.Append(missingItem584);
            pivotCacheRecord32.Append(missingItem585);
            pivotCacheRecord32.Append(missingItem586);
            pivotCacheRecord32.Append(missingItem587);
            pivotCacheRecord32.Append(missingItem588);
            pivotCacheRecord32.Append(missingItem589);
            pivotCacheRecord32.Append(missingItem590);
            pivotCacheRecord32.Append(missingItem591);
            pivotCacheRecord32.Append(fieldItem96);

            PivotCacheRecord pivotCacheRecord33 = new PivotCacheRecord();
            MissingItem missingItem592 = new MissingItem();
            MissingItem missingItem593 = new MissingItem();
            FieldItem fieldItem97 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem594 = new MissingItem();
            MissingItem missingItem595 = new MissingItem();
            FieldItem fieldItem98 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem596 = new MissingItem();
            MissingItem missingItem597 = new MissingItem();
            MissingItem missingItem598 = new MissingItem();
            MissingItem missingItem599 = new MissingItem();
            MissingItem missingItem600 = new MissingItem();
            MissingItem missingItem601 = new MissingItem();
            MissingItem missingItem602 = new MissingItem();
            MissingItem missingItem603 = new MissingItem();
            MissingItem missingItem604 = new MissingItem();
            MissingItem missingItem605 = new MissingItem();
            MissingItem missingItem606 = new MissingItem();
            MissingItem missingItem607 = new MissingItem();
            MissingItem missingItem608 = new MissingItem();
            MissingItem missingItem609 = new MissingItem();
            MissingItem missingItem610 = new MissingItem();
            MissingItem missingItem611 = new MissingItem();
            MissingItem missingItem612 = new MissingItem();
            MissingItem missingItem613 = new MissingItem();
            MissingItem missingItem614 = new MissingItem();
            MissingItem missingItem615 = new MissingItem();
            FieldItem fieldItem99 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord33.Append(missingItem592);
            pivotCacheRecord33.Append(missingItem593);
            pivotCacheRecord33.Append(fieldItem97);
            pivotCacheRecord33.Append(missingItem594);
            pivotCacheRecord33.Append(missingItem595);
            pivotCacheRecord33.Append(fieldItem98);
            pivotCacheRecord33.Append(missingItem596);
            pivotCacheRecord33.Append(missingItem597);
            pivotCacheRecord33.Append(missingItem598);
            pivotCacheRecord33.Append(missingItem599);
            pivotCacheRecord33.Append(missingItem600);
            pivotCacheRecord33.Append(missingItem601);
            pivotCacheRecord33.Append(missingItem602);
            pivotCacheRecord33.Append(missingItem603);
            pivotCacheRecord33.Append(missingItem604);
            pivotCacheRecord33.Append(missingItem605);
            pivotCacheRecord33.Append(missingItem606);
            pivotCacheRecord33.Append(missingItem607);
            pivotCacheRecord33.Append(missingItem608);
            pivotCacheRecord33.Append(missingItem609);
            pivotCacheRecord33.Append(missingItem610);
            pivotCacheRecord33.Append(missingItem611);
            pivotCacheRecord33.Append(missingItem612);
            pivotCacheRecord33.Append(missingItem613);
            pivotCacheRecord33.Append(missingItem614);
            pivotCacheRecord33.Append(missingItem615);
            pivotCacheRecord33.Append(fieldItem99);

            PivotCacheRecord pivotCacheRecord34 = new PivotCacheRecord();
            MissingItem missingItem616 = new MissingItem();
            MissingItem missingItem617 = new MissingItem();
            FieldItem fieldItem100 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem618 = new MissingItem();
            MissingItem missingItem619 = new MissingItem();
            FieldItem fieldItem101 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem620 = new MissingItem();
            MissingItem missingItem621 = new MissingItem();
            MissingItem missingItem622 = new MissingItem();
            MissingItem missingItem623 = new MissingItem();
            MissingItem missingItem624 = new MissingItem();
            MissingItem missingItem625 = new MissingItem();
            MissingItem missingItem626 = new MissingItem();
            MissingItem missingItem627 = new MissingItem();
            MissingItem missingItem628 = new MissingItem();
            MissingItem missingItem629 = new MissingItem();
            MissingItem missingItem630 = new MissingItem();
            MissingItem missingItem631 = new MissingItem();
            MissingItem missingItem632 = new MissingItem();
            MissingItem missingItem633 = new MissingItem();
            MissingItem missingItem634 = new MissingItem();
            MissingItem missingItem635 = new MissingItem();
            MissingItem missingItem636 = new MissingItem();
            MissingItem missingItem637 = new MissingItem();
            MissingItem missingItem638 = new MissingItem();
            MissingItem missingItem639 = new MissingItem();
            FieldItem fieldItem102 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord34.Append(missingItem616);
            pivotCacheRecord34.Append(missingItem617);
            pivotCacheRecord34.Append(fieldItem100);
            pivotCacheRecord34.Append(missingItem618);
            pivotCacheRecord34.Append(missingItem619);
            pivotCacheRecord34.Append(fieldItem101);
            pivotCacheRecord34.Append(missingItem620);
            pivotCacheRecord34.Append(missingItem621);
            pivotCacheRecord34.Append(missingItem622);
            pivotCacheRecord34.Append(missingItem623);
            pivotCacheRecord34.Append(missingItem624);
            pivotCacheRecord34.Append(missingItem625);
            pivotCacheRecord34.Append(missingItem626);
            pivotCacheRecord34.Append(missingItem627);
            pivotCacheRecord34.Append(missingItem628);
            pivotCacheRecord34.Append(missingItem629);
            pivotCacheRecord34.Append(missingItem630);
            pivotCacheRecord34.Append(missingItem631);
            pivotCacheRecord34.Append(missingItem632);
            pivotCacheRecord34.Append(missingItem633);
            pivotCacheRecord34.Append(missingItem634);
            pivotCacheRecord34.Append(missingItem635);
            pivotCacheRecord34.Append(missingItem636);
            pivotCacheRecord34.Append(missingItem637);
            pivotCacheRecord34.Append(missingItem638);
            pivotCacheRecord34.Append(missingItem639);
            pivotCacheRecord34.Append(fieldItem102);

            PivotCacheRecord pivotCacheRecord35 = new PivotCacheRecord();
            MissingItem missingItem640 = new MissingItem();
            MissingItem missingItem641 = new MissingItem();
            FieldItem fieldItem103 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem642 = new MissingItem();
            MissingItem missingItem643 = new MissingItem();
            FieldItem fieldItem104 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem644 = new MissingItem();
            MissingItem missingItem645 = new MissingItem();
            MissingItem missingItem646 = new MissingItem();
            MissingItem missingItem647 = new MissingItem();
            MissingItem missingItem648 = new MissingItem();
            MissingItem missingItem649 = new MissingItem();
            MissingItem missingItem650 = new MissingItem();
            MissingItem missingItem651 = new MissingItem();
            MissingItem missingItem652 = new MissingItem();
            MissingItem missingItem653 = new MissingItem();
            MissingItem missingItem654 = new MissingItem();
            MissingItem missingItem655 = new MissingItem();
            MissingItem missingItem656 = new MissingItem();
            MissingItem missingItem657 = new MissingItem();
            MissingItem missingItem658 = new MissingItem();
            MissingItem missingItem659 = new MissingItem();
            MissingItem missingItem660 = new MissingItem();
            MissingItem missingItem661 = new MissingItem();
            MissingItem missingItem662 = new MissingItem();
            MissingItem missingItem663 = new MissingItem();
            FieldItem fieldItem105 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord35.Append(missingItem640);
            pivotCacheRecord35.Append(missingItem641);
            pivotCacheRecord35.Append(fieldItem103);
            pivotCacheRecord35.Append(missingItem642);
            pivotCacheRecord35.Append(missingItem643);
            pivotCacheRecord35.Append(fieldItem104);
            pivotCacheRecord35.Append(missingItem644);
            pivotCacheRecord35.Append(missingItem645);
            pivotCacheRecord35.Append(missingItem646);
            pivotCacheRecord35.Append(missingItem647);
            pivotCacheRecord35.Append(missingItem648);
            pivotCacheRecord35.Append(missingItem649);
            pivotCacheRecord35.Append(missingItem650);
            pivotCacheRecord35.Append(missingItem651);
            pivotCacheRecord35.Append(missingItem652);
            pivotCacheRecord35.Append(missingItem653);
            pivotCacheRecord35.Append(missingItem654);
            pivotCacheRecord35.Append(missingItem655);
            pivotCacheRecord35.Append(missingItem656);
            pivotCacheRecord35.Append(missingItem657);
            pivotCacheRecord35.Append(missingItem658);
            pivotCacheRecord35.Append(missingItem659);
            pivotCacheRecord35.Append(missingItem660);
            pivotCacheRecord35.Append(missingItem661);
            pivotCacheRecord35.Append(missingItem662);
            pivotCacheRecord35.Append(missingItem663);
            pivotCacheRecord35.Append(fieldItem105);

            PivotCacheRecord pivotCacheRecord36 = new PivotCacheRecord();
            MissingItem missingItem664 = new MissingItem();
            MissingItem missingItem665 = new MissingItem();
            FieldItem fieldItem106 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem666 = new MissingItem();
            MissingItem missingItem667 = new MissingItem();
            FieldItem fieldItem107 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem668 = new MissingItem();
            MissingItem missingItem669 = new MissingItem();
            MissingItem missingItem670 = new MissingItem();
            MissingItem missingItem671 = new MissingItem();
            MissingItem missingItem672 = new MissingItem();
            MissingItem missingItem673 = new MissingItem();
            MissingItem missingItem674 = new MissingItem();
            MissingItem missingItem675 = new MissingItem();
            MissingItem missingItem676 = new MissingItem();
            MissingItem missingItem677 = new MissingItem();
            MissingItem missingItem678 = new MissingItem();
            MissingItem missingItem679 = new MissingItem();
            MissingItem missingItem680 = new MissingItem();
            MissingItem missingItem681 = new MissingItem();
            MissingItem missingItem682 = new MissingItem();
            MissingItem missingItem683 = new MissingItem();
            MissingItem missingItem684 = new MissingItem();
            MissingItem missingItem685 = new MissingItem();
            MissingItem missingItem686 = new MissingItem();
            MissingItem missingItem687 = new MissingItem();
            FieldItem fieldItem108 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord36.Append(missingItem664);
            pivotCacheRecord36.Append(missingItem665);
            pivotCacheRecord36.Append(fieldItem106);
            pivotCacheRecord36.Append(missingItem666);
            pivotCacheRecord36.Append(missingItem667);
            pivotCacheRecord36.Append(fieldItem107);
            pivotCacheRecord36.Append(missingItem668);
            pivotCacheRecord36.Append(missingItem669);
            pivotCacheRecord36.Append(missingItem670);
            pivotCacheRecord36.Append(missingItem671);
            pivotCacheRecord36.Append(missingItem672);
            pivotCacheRecord36.Append(missingItem673);
            pivotCacheRecord36.Append(missingItem674);
            pivotCacheRecord36.Append(missingItem675);
            pivotCacheRecord36.Append(missingItem676);
            pivotCacheRecord36.Append(missingItem677);
            pivotCacheRecord36.Append(missingItem678);
            pivotCacheRecord36.Append(missingItem679);
            pivotCacheRecord36.Append(missingItem680);
            pivotCacheRecord36.Append(missingItem681);
            pivotCacheRecord36.Append(missingItem682);
            pivotCacheRecord36.Append(missingItem683);
            pivotCacheRecord36.Append(missingItem684);
            pivotCacheRecord36.Append(missingItem685);
            pivotCacheRecord36.Append(missingItem686);
            pivotCacheRecord36.Append(missingItem687);
            pivotCacheRecord36.Append(fieldItem108);

            PivotCacheRecord pivotCacheRecord37 = new PivotCacheRecord();
            MissingItem missingItem688 = new MissingItem();
            MissingItem missingItem689 = new MissingItem();
            FieldItem fieldItem109 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem690 = new MissingItem();
            MissingItem missingItem691 = new MissingItem();
            FieldItem fieldItem110 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem692 = new MissingItem();
            MissingItem missingItem693 = new MissingItem();
            MissingItem missingItem694 = new MissingItem();
            MissingItem missingItem695 = new MissingItem();
            MissingItem missingItem696 = new MissingItem();
            MissingItem missingItem697 = new MissingItem();
            MissingItem missingItem698 = new MissingItem();
            MissingItem missingItem699 = new MissingItem();
            MissingItem missingItem700 = new MissingItem();
            MissingItem missingItem701 = new MissingItem();
            MissingItem missingItem702 = new MissingItem();
            MissingItem missingItem703 = new MissingItem();
            MissingItem missingItem704 = new MissingItem();
            MissingItem missingItem705 = new MissingItem();
            MissingItem missingItem706 = new MissingItem();
            MissingItem missingItem707 = new MissingItem();
            MissingItem missingItem708 = new MissingItem();
            MissingItem missingItem709 = new MissingItem();
            MissingItem missingItem710 = new MissingItem();
            MissingItem missingItem711 = new MissingItem();
            FieldItem fieldItem111 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord37.Append(missingItem688);
            pivotCacheRecord37.Append(missingItem689);
            pivotCacheRecord37.Append(fieldItem109);
            pivotCacheRecord37.Append(missingItem690);
            pivotCacheRecord37.Append(missingItem691);
            pivotCacheRecord37.Append(fieldItem110);
            pivotCacheRecord37.Append(missingItem692);
            pivotCacheRecord37.Append(missingItem693);
            pivotCacheRecord37.Append(missingItem694);
            pivotCacheRecord37.Append(missingItem695);
            pivotCacheRecord37.Append(missingItem696);
            pivotCacheRecord37.Append(missingItem697);
            pivotCacheRecord37.Append(missingItem698);
            pivotCacheRecord37.Append(missingItem699);
            pivotCacheRecord37.Append(missingItem700);
            pivotCacheRecord37.Append(missingItem701);
            pivotCacheRecord37.Append(missingItem702);
            pivotCacheRecord37.Append(missingItem703);
            pivotCacheRecord37.Append(missingItem704);
            pivotCacheRecord37.Append(missingItem705);
            pivotCacheRecord37.Append(missingItem706);
            pivotCacheRecord37.Append(missingItem707);
            pivotCacheRecord37.Append(missingItem708);
            pivotCacheRecord37.Append(missingItem709);
            pivotCacheRecord37.Append(missingItem710);
            pivotCacheRecord37.Append(missingItem711);
            pivotCacheRecord37.Append(fieldItem111);

            PivotCacheRecord pivotCacheRecord38 = new PivotCacheRecord();
            MissingItem missingItem712 = new MissingItem();
            MissingItem missingItem713 = new MissingItem();
            FieldItem fieldItem112 = new FieldItem(){ Val = (UInt32Value)8U };
            MissingItem missingItem714 = new MissingItem();
            MissingItem missingItem715 = new MissingItem();
            FieldItem fieldItem113 = new FieldItem(){ Val = (UInt32Value)2U };
            MissingItem missingItem716 = new MissingItem();
            MissingItem missingItem717 = new MissingItem();
            MissingItem missingItem718 = new MissingItem();
            MissingItem missingItem719 = new MissingItem();
            MissingItem missingItem720 = new MissingItem();
            MissingItem missingItem721 = new MissingItem();
            MissingItem missingItem722 = new MissingItem();
            MissingItem missingItem723 = new MissingItem();
            MissingItem missingItem724 = new MissingItem();
            MissingItem missingItem725 = new MissingItem();
            MissingItem missingItem726 = new MissingItem();
            MissingItem missingItem727 = new MissingItem();
            MissingItem missingItem728 = new MissingItem();
            MissingItem missingItem729 = new MissingItem();
            MissingItem missingItem730 = new MissingItem();
            MissingItem missingItem731 = new MissingItem();
            MissingItem missingItem732 = new MissingItem();
            MissingItem missingItem733 = new MissingItem();
            MissingItem missingItem734 = new MissingItem();
            MissingItem missingItem735 = new MissingItem();
            FieldItem fieldItem114 = new FieldItem(){ Val = (UInt32Value)1U };

            pivotCacheRecord38.Append(missingItem712);
            pivotCacheRecord38.Append(missingItem713);
            pivotCacheRecord38.Append(fieldItem112);
            pivotCacheRecord38.Append(missingItem714);
            pivotCacheRecord38.Append(missingItem715);
            pivotCacheRecord38.Append(fieldItem113);
            pivotCacheRecord38.Append(missingItem716);
            pivotCacheRecord38.Append(missingItem717);
            pivotCacheRecord38.Append(missingItem718);
            pivotCacheRecord38.Append(missingItem719);
            pivotCacheRecord38.Append(missingItem720);
            pivotCacheRecord38.Append(missingItem721);
            pivotCacheRecord38.Append(missingItem722);
            pivotCacheRecord38.Append(missingItem723);
            pivotCacheRecord38.Append(missingItem724);
            pivotCacheRecord38.Append(missingItem725);
            pivotCacheRecord38.Append(missingItem726);
            pivotCacheRecord38.Append(missingItem727);
            pivotCacheRecord38.Append(missingItem728);
            pivotCacheRecord38.Append(missingItem729);
            pivotCacheRecord38.Append(missingItem730);
            pivotCacheRecord38.Append(missingItem731);
            pivotCacheRecord38.Append(missingItem732);
            pivotCacheRecord38.Append(missingItem733);
            pivotCacheRecord38.Append(missingItem734);
            pivotCacheRecord38.Append(missingItem735);
            pivotCacheRecord38.Append(fieldItem114);

            pivotCacheRecords1.Append(pivotCacheRecord1);
            pivotCacheRecords1.Append(pivotCacheRecord2);
            pivotCacheRecords1.Append(pivotCacheRecord3);
            pivotCacheRecords1.Append(pivotCacheRecord4);
            pivotCacheRecords1.Append(pivotCacheRecord5);
            pivotCacheRecords1.Append(pivotCacheRecord6);
            pivotCacheRecords1.Append(pivotCacheRecord7);
            pivotCacheRecords1.Append(pivotCacheRecord8);
            pivotCacheRecords1.Append(pivotCacheRecord9);
            pivotCacheRecords1.Append(pivotCacheRecord10);
            pivotCacheRecords1.Append(pivotCacheRecord11);
            pivotCacheRecords1.Append(pivotCacheRecord12);
            pivotCacheRecords1.Append(pivotCacheRecord13);
            pivotCacheRecords1.Append(pivotCacheRecord14);
            pivotCacheRecords1.Append(pivotCacheRecord15);
            pivotCacheRecords1.Append(pivotCacheRecord16);
            pivotCacheRecords1.Append(pivotCacheRecord17);
            pivotCacheRecords1.Append(pivotCacheRecord18);
            pivotCacheRecords1.Append(pivotCacheRecord19);
            pivotCacheRecords1.Append(pivotCacheRecord20);
            pivotCacheRecords1.Append(pivotCacheRecord21);
            pivotCacheRecords1.Append(pivotCacheRecord22);
            pivotCacheRecords1.Append(pivotCacheRecord23);
            pivotCacheRecords1.Append(pivotCacheRecord24);
            pivotCacheRecords1.Append(pivotCacheRecord25);
            pivotCacheRecords1.Append(pivotCacheRecord26);
            pivotCacheRecords1.Append(pivotCacheRecord27);
            pivotCacheRecords1.Append(pivotCacheRecord28);
            pivotCacheRecords1.Append(pivotCacheRecord29);
            pivotCacheRecords1.Append(pivotCacheRecord30);
            pivotCacheRecords1.Append(pivotCacheRecord31);
            pivotCacheRecords1.Append(pivotCacheRecord32);
            pivotCacheRecords1.Append(pivotCacheRecord33);
            pivotCacheRecords1.Append(pivotCacheRecord34);
            pivotCacheRecords1.Append(pivotCacheRecord35);
            pivotCacheRecords1.Append(pivotCacheRecord36);
            pivotCacheRecords1.Append(pivotCacheRecord37);
            pivotCacheRecords1.Append(pivotCacheRecord38);

            pivotTableCacheRecordsPart1.PivotCacheRecords = pivotCacheRecords1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "x14ac xr xr2 xr3" }  };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            worksheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            worksheet1.AddNamespaceDeclaration("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2");
            worksheet1.AddNamespaceDeclaration("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3");
            worksheet1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{00000000-0001-0000-0000-000000000000}"));
            SheetDimension sheetDimension1 = new SheetDimension(){ Reference = "A1:AA32" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView(){ TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection(){ ActiveCell = "E12", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "E12" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties(){ DefaultColumnWidth = 9D, DefaultRowHeight = 13.5D };

            Columns columns1 = new Columns();
            Column column1 = new Column(){ Min = (UInt32Value)5U, Max = (UInt32Value)5U, Width = 13.625D, BestFit = true, CustomWidth = true };
            Column column2 = new Column(){ Min = (UInt32Value)6U, Max = (UInt32Value)6U, Width = 8D, BestFit = true, CustomWidth = true };
            Column column3 = new Column(){ Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 13.625D, BestFit = true, CustomWidth = true };
            Column column4 = new Column(){ Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 9.125D, BestFit = true, CustomWidth = true };
            Column column5 = new Column(){ Min = (UInt32Value)9U, Max = (UInt32Value)9U, Width = 13.625D, BestFit = true, CustomWidth = true };
            Column column6 = new Column(){ Min = (UInt32Value)10U, Max = (UInt32Value)10U, Width = 10.25D, BestFit = true, CustomWidth = true };
            Column column7 = new Column(){ Min = (UInt32Value)11U, Max = (UInt32Value)11U, Width = 9.125D, BestFit = true, CustomWidth = true };
            Column column8 = new Column(){ Min = (UInt32Value)12U, Max = (UInt32Value)14U, Width = 13.625D, BestFit = true, CustomWidth = true };
            Column column9 = new Column(){ Min = (UInt32Value)17U, Max = (UInt32Value)17U, Width = 9.375D };
            Column column10 = new Column(){ Min = (UInt32Value)22U, Max = (UInt32Value)22U, Width = 9.375D };
            Column column11 = new Column(){ Min = (UInt32Value)23U, Max = (UInt32Value)24U, Width = 11.5D };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);
            columns1.Append(column7);
            columns1.Append(column8);
            columns1.Append(column9);
            columns1.Append(column10);
            columns1.Append(column11);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row(){ RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:27" }, StyleIndex = (UInt32Value)1U, CustomFormat = true };

            Cell cell1 = new Cell(){ CellReference = "A1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);

            Cell cell2 = new Cell(){ CellReference = "B1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell2.Append(cellValue2);

            Cell cell3 = new Cell(){ CellReference = "C1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "2";

            cell3.Append(cellValue3);

            Cell cell4 = new Cell(){ CellReference = "D1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "3";

            cell4.Append(cellValue4);

            Cell cell5 = new Cell(){ CellReference = "E1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "4";

            cell5.Append(cellValue5);

            Cell cell6 = new Cell(){ CellReference = "F1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "5";

            cell6.Append(cellValue6);

            Cell cell7 = new Cell(){ CellReference = "G1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "6";

            cell7.Append(cellValue7);

            Cell cell8 = new Cell(){ CellReference = "H1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "7";

            cell8.Append(cellValue8);

            Cell cell9 = new Cell(){ CellReference = "I1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "8";

            cell9.Append(cellValue9);

            Cell cell10 = new Cell(){ CellReference = "J1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "9";

            cell10.Append(cellValue10);

            Cell cell11 = new Cell(){ CellReference = "K1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "10";

            cell11.Append(cellValue11);

            Cell cell12 = new Cell(){ CellReference = "L1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "11";

            cell12.Append(cellValue12);

            Cell cell13 = new Cell(){ CellReference = "M1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "12";

            cell13.Append(cellValue13);

            Cell cell14 = new Cell(){ CellReference = "N1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "13";

            cell14.Append(cellValue14);

            Cell cell15 = new Cell(){ CellReference = "O1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "14";

            cell15.Append(cellValue15);

            Cell cell16 = new Cell(){ CellReference = "P1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "15";

            cell16.Append(cellValue16);

            Cell cell17 = new Cell(){ CellReference = "Q1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue17 = new CellValue();
            cellValue17.Text = "16";

            cell17.Append(cellValue17);

            Cell cell18 = new Cell(){ CellReference = "R1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue18 = new CellValue();
            cellValue18.Text = "17";

            cell18.Append(cellValue18);

            Cell cell19 = new Cell(){ CellReference = "S1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue19 = new CellValue();
            cellValue19.Text = "18";

            cell19.Append(cellValue19);

            Cell cell20 = new Cell(){ CellReference = "T1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = "19";

            cell20.Append(cellValue20);

            Cell cell21 = new Cell(){ CellReference = "U1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "20";

            cell21.Append(cellValue21);

            Cell cell22 = new Cell(){ CellReference = "V1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "21";

            cell22.Append(cellValue22);

            Cell cell23 = new Cell(){ CellReference = "W1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "22";

            cell23.Append(cellValue23);

            Cell cell24 = new Cell(){ CellReference = "X1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue24 = new CellValue();
            cellValue24.Text = "23";

            cell24.Append(cellValue24);

            Cell cell25 = new Cell(){ CellReference = "Y1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue25 = new CellValue();
            cellValue25.Text = "24";

            cell25.Append(cellValue25);

            Cell cell26 = new Cell(){ CellReference = "Z1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue26 = new CellValue();
            cellValue26.Text = "25";

            cell26.Append(cellValue26);

            Cell cell27 = new Cell(){ CellReference = "AA1", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue27 = new CellValue();
            cellValue27.Text = "26";

            cell27.Append(cellValue27);

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);
            row1.Append(cell7);
            row1.Append(cell8);
            row1.Append(cell9);
            row1.Append(cell10);
            row1.Append(cell11);
            row1.Append(cell12);
            row1.Append(cell13);
            row1.Append(cell14);
            row1.Append(cell15);
            row1.Append(cell16);
            row1.Append(cell17);
            row1.Append(cell18);
            row1.Append(cell19);
            row1.Append(cell20);
            row1.Append(cell21);
            row1.Append(cell22);
            row1.Append(cell23);
            row1.Append(cell24);
            row1.Append(cell25);
            row1.Append(cell26);
            row1.Append(cell27);

            Row row2 = new Row(){ RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell28 = new Cell(){ CellReference = "A2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "27";

            cell28.Append(cellValue28);

            Cell cell29 = new Cell(){ CellReference = "B2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "28";

            cell29.Append(cellValue29);

            Cell cell30 = new Cell(){ CellReference = "C2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "27";

            cell30.Append(cellValue30);

            Cell cell31 = new Cell(){ CellReference = "D2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "29";

            cell31.Append(cellValue31);

            Cell cell32 = new Cell(){ CellReference = "E2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "30";

            cell32.Append(cellValue32);

            Cell cell33 = new Cell(){ CellReference = "F2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue33 = new CellValue();
            cellValue33.Text = "31";

            cell33.Append(cellValue33);

            Cell cell34 = new Cell(){ CellReference = "G2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue34 = new CellValue();
            cellValue34.Text = "0";

            cell34.Append(cellValue34);

            Cell cell35 = new Cell(){ CellReference = "H2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue35 = new CellValue();
            cellValue35.Text = "12.5";

            cell35.Append(cellValue35);

            Cell cell36 = new Cell(){ CellReference = "I2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue36 = new CellValue();
            cellValue36.Text = "17";

            cell36.Append(cellValue36);

            Cell cell37 = new Cell(){ CellReference = "J2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = "212.47";

            cell37.Append(cellValue37);

            Cell cell38 = new Cell(){ CellReference = "K2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "1604.9";

            cell38.Append(cellValue38);

            Cell cell39 = new Cell(){ CellReference = "L2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "8024.5";

            cell39.Append(cellValue39);

            Cell cell40 = new Cell(){ CellReference = "M2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "67.56";

            cell40.Append(cellValue40);

            Cell cell41 = new Cell(){ CellReference = "N2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue41 = new CellValue();
            cellValue41.Text = "579.73400000000004";

            cell41.Append(cellValue41);

            Cell cell42 = new Cell(){ CellReference = "O2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue42 = new CellValue();
            cellValue42.Text = "2898.67";

            cell42.Append(cellValue42);

            Cell cell43 = new Cell(){ CellReference = "P2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue43 = new CellValue();
            cellValue43.Text = "0.318";

            cell43.Append(cellValue43);

            Cell cell44 = new Cell(){ CellReference = "Q2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue44 = new CellValue();
            cellValue44.Text = "0.36122700000000002";

            cell44.Append(cellValue44);

            Cell cell45 = new Cell(){ CellReference = "R2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue45 = new CellValue();
            cellValue45.Text = "2.99";

            cell45.Append(cellValue45);

            Cell cell46 = new Cell(){ CellReference = "S2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue46 = new CellValue();
            cellValue46.Text = "36.177999999999997";

            cell46.Append(cellValue46);

            Cell cell47 = new Cell(){ CellReference = "T2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue47 = new CellValue();
            cellValue47.Text = "180.89";

            cell47.Append(cellValue47);

            Cell cell48 = new Cell(){ CellReference = "U2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue48 = new CellValue();
            cellValue48.Text = "1.41E-2";

            cell48.Append(cellValue48);

            Cell cell49 = new Cell(){ CellReference = "V2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue49 = new CellValue();
            cellValue49.Text = "2.2542E-2";

            cell49.Append(cellValue49);

            Cell cell50 = new Cell(){ CellReference = "W2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue50 = new CellValue();
            cellValue50.Text = "1.4769126699999999";

            cell50.Append(cellValue50);

            Cell cell51 = new Cell(){ CellReference = "X2", StyleIndex = (UInt32Value)2U };
            CellValue cellValue51 = new CellValue();
            cellValue51.Text = "0.13130290999999999";

            cell51.Append(cellValue51);
            Cell cell52 = new Cell(){ CellReference = "Y2", StyleIndex = (UInt32Value)2U };
            Cell cell53 = new Cell(){ CellReference = "Z2", StyleIndex = (UInt32Value)2U };

            Cell cell54 = new Cell(){ CellReference = "AA2", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue52 = new CellValue();
            cellValue52.Text = "32";

            cell54.Append(cellValue52);

            row2.Append(cell28);
            row2.Append(cell29);
            row2.Append(cell30);
            row2.Append(cell31);
            row2.Append(cell32);
            row2.Append(cell33);
            row2.Append(cell34);
            row2.Append(cell35);
            row2.Append(cell36);
            row2.Append(cell37);
            row2.Append(cell38);
            row2.Append(cell39);
            row2.Append(cell40);
            row2.Append(cell41);
            row2.Append(cell42);
            row2.Append(cell43);
            row2.Append(cell44);
            row2.Append(cell45);
            row2.Append(cell46);
            row2.Append(cell47);
            row2.Append(cell48);
            row2.Append(cell49);
            row2.Append(cell50);
            row2.Append(cell51);
            row2.Append(cell52);
            row2.Append(cell53);
            row2.Append(cell54);

            Row row3 = new Row(){ RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell55 = new Cell(){ CellReference = "A3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue53 = new CellValue();
            cellValue53.Text = "33";

            cell55.Append(cellValue53);

            Cell cell56 = new Cell(){ CellReference = "B3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue54 = new CellValue();
            cellValue54.Text = "28";

            cell56.Append(cellValue54);

            Cell cell57 = new Cell(){ CellReference = "C3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue55 = new CellValue();
            cellValue55.Text = "34";

            cell57.Append(cellValue55);

            Cell cell58 = new Cell(){ CellReference = "D3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue56 = new CellValue();
            cellValue56.Text = "35";

            cell58.Append(cellValue56);

            Cell cell59 = new Cell(){ CellReference = "E3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue57 = new CellValue();
            cellValue57.Text = "36";

            cell59.Append(cellValue57);

            Cell cell60 = new Cell(){ CellReference = "F3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue58 = new CellValue();
            cellValue58.Text = "37";

            cell60.Append(cellValue58);

            Cell cell61 = new Cell(){ CellReference = "G3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue59 = new CellValue();
            cellValue59.Text = "0";

            cell61.Append(cellValue59);

            Cell cell62 = new Cell(){ CellReference = "H3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue60 = new CellValue();
            cellValue60.Text = "18.690000000000001";

            cell62.Append(cellValue60);

            Cell cell63 = new Cell(){ CellReference = "I3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue61 = new CellValue();
            cellValue61.Text = "4";

            cell63.Append(cellValue61);

            Cell cell64 = new Cell(){ CellReference = "J3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue62 = new CellValue();
            cellValue62.Text = "74.760000000000005";

            cell64.Append(cellValue62);

            Cell cell65 = new Cell(){ CellReference = "K3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue63 = new CellValue();
            cellValue63.Text = "266.79000000000002";

            cell65.Append(cellValue63);

            Cell cell66 = new Cell(){ CellReference = "L3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue64 = new CellValue();
            cellValue64.Text = "1067.1600000000001";

            cell66.Append(cellValue64);

            Cell cell67 = new Cell(){ CellReference = "M3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue65 = new CellValue();
            cellValue65.Text = "33.72";

            cell67.Append(cellValue65);

            Cell cell68 = new Cell(){ CellReference = "N3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue66 = new CellValue();
            cellValue66.Text = "117.83";

            cell68.Append(cellValue66);

            Cell cell69 = new Cell(){ CellReference = "O3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue67 = new CellValue();
            cellValue67.Text = "471.32";

            cell69.Append(cellValue67);

            Cell cell70 = new Cell(){ CellReference = "P3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue68 = new CellValue();
            cellValue68.Text = "0.45100000000000001";

            cell70.Append(cellValue68);

            Cell cell71 = new Cell(){ CellReference = "Q3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue69 = new CellValue();
            cellValue69.Text = "0.441658";

            cell71.Append(cellValue69);

            Cell cell72 = new Cell(){ CellReference = "R3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue70 = new CellValue();
            cellValue70.Text = "0";

            cell72.Append(cellValue70);

            Cell cell73 = new Cell(){ CellReference = "S3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue71 = new CellValue();
            cellValue71.Text = "23.487500000000001";

            cell73.Append(cellValue71);

            Cell cell74 = new Cell(){ CellReference = "T3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue72 = new CellValue();
            cellValue72.Text = "93.95";

            cell74.Append(cellValue72);

            Cell cell75 = new Cell(){ CellReference = "U3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue73 = new CellValue();
            cellValue73.Text = "0";

            cell75.Append(cellValue73);

            Cell cell76 = new Cell(){ CellReference = "V3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue74 = new CellValue();
            cellValue74.Text = "8.8037000000000004E-2";

            cell76.Append(cellValue74);

            Cell cell77 = new Cell(){ CellReference = "W3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue75 = new CellValue();
            cellValue75.Text = "0.48286602000000001";

            cell77.Append(cellValue75);

            Cell cell78 = new Cell(){ CellReference = "X3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue76 = new CellValue();
            cellValue76.Text = "-0.43385080999999998";

            cell78.Append(cellValue76);

            Cell cell79 = new Cell(){ CellReference = "Y3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue77 = new CellValue();
            cellValue77.Text = "-0.50734760999999995";

            cell79.Append(cellValue77);

            Cell cell80 = new Cell(){ CellReference = "Z3", StyleIndex = (UInt32Value)2U };
            CellValue cellValue78 = new CellValue();
            cellValue78.Text = "0.77675992000000005";

            cell80.Append(cellValue78);

            Cell cell81 = new Cell(){ CellReference = "AA3", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue79 = new CellValue();
            cellValue79.Text = "32";

            cell81.Append(cellValue79);

            row3.Append(cell55);
            row3.Append(cell56);
            row3.Append(cell57);
            row3.Append(cell58);
            row3.Append(cell59);
            row3.Append(cell60);
            row3.Append(cell61);
            row3.Append(cell62);
            row3.Append(cell63);
            row3.Append(cell64);
            row3.Append(cell65);
            row3.Append(cell66);
            row3.Append(cell67);
            row3.Append(cell68);
            row3.Append(cell69);
            row3.Append(cell70);
            row3.Append(cell71);
            row3.Append(cell72);
            row3.Append(cell73);
            row3.Append(cell74);
            row3.Append(cell75);
            row3.Append(cell76);
            row3.Append(cell77);
            row3.Append(cell78);
            row3.Append(cell79);
            row3.Append(cell80);
            row3.Append(cell81);

            Row row4 = new Row(){ RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell82 = new Cell(){ CellReference = "A4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue80 = new CellValue();
            cellValue80.Text = "38";

            cell82.Append(cellValue80);

            Cell cell83 = new Cell(){ CellReference = "B4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue81 = new CellValue();
            cellValue81.Text = "39";

            cell83.Append(cellValue81);

            Cell cell84 = new Cell(){ CellReference = "C4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue82 = new CellValue();
            cellValue82.Text = "40";

            cell84.Append(cellValue82);

            Cell cell85 = new Cell(){ CellReference = "D4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue83 = new CellValue();
            cellValue83.Text = "41";

            cell85.Append(cellValue83);

            Cell cell86 = new Cell(){ CellReference = "E4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue84 = new CellValue();
            cellValue84.Text = "42";

            cell86.Append(cellValue84);

            Cell cell87 = new Cell(){ CellReference = "F4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue85 = new CellValue();
            cellValue85.Text = "31";

            cell87.Append(cellValue85);

            Cell cell88 = new Cell(){ CellReference = "G4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue86 = new CellValue();
            cellValue86.Text = "0";

            cell88.Append(cellValue86);

            Cell cell89 = new Cell(){ CellReference = "H4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue87 = new CellValue();
            cellValue87.Text = "12.28";

            cell89.Append(cellValue87);

            Cell cell90 = new Cell(){ CellReference = "I4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue88 = new CellValue();
            cellValue88.Text = "58";

            cell90.Append(cellValue88);

            Cell cell91 = new Cell(){ CellReference = "J4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue89 = new CellValue();
            cellValue89.Text = "712.48";

            cell91.Append(cellValue89);

            Cell cell92 = new Cell(){ CellReference = "K4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue90 = new CellValue();
            cellValue90.Text = "830.2";

            cell92.Append(cellValue90);

            Cell cell93 = new Cell(){ CellReference = "L4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue91 = new CellValue();
            cellValue91.Text = "1660.4";

            cell93.Append(cellValue91);

            Cell cell94 = new Cell(){ CellReference = "M4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue92 = new CellValue();
            cellValue92.Text = "389.25";

            cell94.Append(cellValue92);

            Cell cell95 = new Cell(){ CellReference = "N4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue93 = new CellValue();
            cellValue93.Text = "408.71";

            cell95.Append(cellValue93);

            Cell cell96 = new Cell(){ CellReference = "O4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue94 = new CellValue();
            cellValue94.Text = "817.42";

            cell96.Append(cellValue94);

            Cell cell97 = new Cell(){ CellReference = "P4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue95 = new CellValue();
            cellValue95.Text = "0.54630000000000001";

            cell97.Append(cellValue95);

            Cell cell98 = new Cell(){ CellReference = "Q4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue96 = new CellValue();
            cellValue96.Text = "0.49230299999999999";

            cell98.Append(cellValue96);

            Cell cell99 = new Cell(){ CellReference = "R4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue97 = new CellValue();
            cellValue97.Text = "0";

            cell99.Append(cellValue97);

            Cell cell100 = new Cell(){ CellReference = "S4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue98 = new CellValue();
            cellValue98.Text = "14.45";

            cell100.Append(cellValue98);

            Cell cell101 = new Cell(){ CellReference = "T4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue99 = new CellValue();
            cellValue99.Text = "28.9";

            cell101.Append(cellValue99);

            Cell cell102 = new Cell(){ CellReference = "U4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue100 = new CellValue();
            cellValue100.Text = "0";

            cell102.Append(cellValue100);

            Cell cell103 = new Cell(){ CellReference = "V4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue101 = new CellValue();
            cellValue101.Text = "1.7405E-2";

            cell103.Append(cellValue101);

            Cell cell104 = new Cell(){ CellReference = "W4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue102 = new CellValue();
            cellValue102.Text = "1.60369742";

            cell104.Append(cellValue102);

            Cell cell105 = new Cell(){ CellReference = "X4", StyleIndex = (UInt32Value)2U };
            CellValue cellValue103 = new CellValue();
            cellValue103.Text = "-7.2327900000000001E-2";

            cell105.Append(cellValue103);
            Cell cell106 = new Cell(){ CellReference = "Y4", StyleIndex = (UInt32Value)2U };
            Cell cell107 = new Cell(){ CellReference = "Z4", StyleIndex = (UInt32Value)2U };

            Cell cell108 = new Cell(){ CellReference = "AA4", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue104 = new CellValue();
            cellValue104.Text = "32";

            cell108.Append(cellValue104);

            row4.Append(cell82);
            row4.Append(cell83);
            row4.Append(cell84);
            row4.Append(cell85);
            row4.Append(cell86);
            row4.Append(cell87);
            row4.Append(cell88);
            row4.Append(cell89);
            row4.Append(cell90);
            row4.Append(cell91);
            row4.Append(cell92);
            row4.Append(cell93);
            row4.Append(cell94);
            row4.Append(cell95);
            row4.Append(cell96);
            row4.Append(cell97);
            row4.Append(cell98);
            row4.Append(cell99);
            row4.Append(cell100);
            row4.Append(cell101);
            row4.Append(cell102);
            row4.Append(cell103);
            row4.Append(cell104);
            row4.Append(cell105);
            row4.Append(cell106);
            row4.Append(cell107);
            row4.Append(cell108);

            Row row5 = new Row(){ RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell109 = new Cell(){ CellReference = "A5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue105 = new CellValue();
            cellValue105.Text = "43";

            cell109.Append(cellValue105);

            Cell cell110 = new Cell(){ CellReference = "B5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue106 = new CellValue();
            cellValue106.Text = "44";

            cell110.Append(cellValue106);

            Cell cell111 = new Cell(){ CellReference = "C5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue107 = new CellValue();
            cellValue107.Text = "45";

            cell111.Append(cellValue107);

            Cell cell112 = new Cell(){ CellReference = "D5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue108 = new CellValue();
            cellValue108.Text = "46";

            cell112.Append(cellValue108);

            Cell cell113 = new Cell(){ CellReference = "E5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue109 = new CellValue();
            cellValue109.Text = "47";

            cell113.Append(cellValue109);

            Cell cell114 = new Cell(){ CellReference = "F5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue110 = new CellValue();
            cellValue110.Text = "37";

            cell114.Append(cellValue110);

            Cell cell115 = new Cell(){ CellReference = "G5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue111 = new CellValue();
            cellValue111.Text = "0";

            cell115.Append(cellValue111);

            Cell cell116 = new Cell(){ CellReference = "H5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue112 = new CellValue();
            cellValue112.Text = "10.78";

            cell116.Append(cellValue112);

            Cell cell117 = new Cell(){ CellReference = "I5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue113 = new CellValue();
            cellValue113.Text = "42";

            cell117.Append(cellValue113);

            Cell cell118 = new Cell(){ CellReference = "J5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue114 = new CellValue();
            cellValue114.Text = "452.84";

            cell118.Append(cellValue114);

            Cell cell119 = new Cell(){ CellReference = "K5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue115 = new CellValue();
            cellValue115.Text = "798.62";

            cell119.Append(cellValue115);

            Cell cell120 = new Cell(){ CellReference = "L5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue116 = new CellValue();
            cellValue116.Text = "3194.48";

            cell120.Append(cellValue116);

            Cell cell121 = new Cell(){ CellReference = "M5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue117 = new CellValue();
            cellValue117.Text = "231.77";

            cell121.Append(cellValue117);

            Cell cell122 = new Cell(){ CellReference = "N5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue118 = new CellValue();
            cellValue118.Text = "374.93";

            cell122.Append(cellValue118);

            Cell cell123 = new Cell(){ CellReference = "O5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue119 = new CellValue();
            cellValue119.Text = "1499.72";

            cell123.Append(cellValue119);

            Cell cell124 = new Cell(){ CellReference = "P5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue120 = new CellValue();
            cellValue120.Text = "0.51180000000000003";

            cell124.Append(cellValue120);

            Cell cell125 = new Cell(){ CellReference = "Q5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue121 = new CellValue();
            cellValue121.Text = "0.469472";

            cell125.Append(cellValue121);

            Cell cell126 = new Cell(){ CellReference = "R5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue122 = new CellValue();
            cellValue122.Text = "18.899999999999999";

            cell126.Append(cellValue122);

            Cell cell127 = new Cell(){ CellReference = "S5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue123 = new CellValue();
            cellValue123.Text = "33.207500000000003";

            cell127.Append(cellValue123);

            Cell cell128 = new Cell(){ CellReference = "T5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue124 = new CellValue();
            cellValue124.Text = "132.83000000000001";

            cell128.Append(cellValue124);

            Cell cell129 = new Cell(){ CellReference = "U5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue125 = new CellValue();
            cellValue125.Text = "4.1700000000000001E-2";

            cell129.Append(cellValue125);

            Cell cell130 = new Cell(){ CellReference = "V5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue126 = new CellValue();
            cellValue126.Text = "4.1581E-2";

            cell130.Append(cellValue126);

            Cell cell131 = new Cell(){ CellReference = "W5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue127 = new CellValue();
            cellValue127.Text = "1.50967355";

            cell131.Append(cellValue127);

            Cell cell132 = new Cell(){ CellReference = "X5", StyleIndex = (UInt32Value)2U };
            CellValue cellValue128 = new CellValue();
            cellValue128.Text = "0.15724105999999999";

            cell132.Append(cellValue128);
            Cell cell133 = new Cell(){ CellReference = "Y5", StyleIndex = (UInt32Value)2U };
            Cell cell134 = new Cell(){ CellReference = "Z5", StyleIndex = (UInt32Value)2U };

            Cell cell135 = new Cell(){ CellReference = "AA5", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue129 = new CellValue();
            cellValue129.Text = "32";

            cell135.Append(cellValue129);

            row5.Append(cell109);
            row5.Append(cell110);
            row5.Append(cell111);
            row5.Append(cell112);
            row5.Append(cell113);
            row5.Append(cell114);
            row5.Append(cell115);
            row5.Append(cell116);
            row5.Append(cell117);
            row5.Append(cell118);
            row5.Append(cell119);
            row5.Append(cell120);
            row5.Append(cell121);
            row5.Append(cell122);
            row5.Append(cell123);
            row5.Append(cell124);
            row5.Append(cell125);
            row5.Append(cell126);
            row5.Append(cell127);
            row5.Append(cell128);
            row5.Append(cell129);
            row5.Append(cell130);
            row5.Append(cell131);
            row5.Append(cell132);
            row5.Append(cell133);
            row5.Append(cell134);
            row5.Append(cell135);

            Row row6 = new Row(){ RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell136 = new Cell(){ CellReference = "A6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue130 = new CellValue();
            cellValue130.Text = "48";

            cell136.Append(cellValue130);

            Cell cell137 = new Cell(){ CellReference = "B6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue131 = new CellValue();
            cellValue131.Text = "49";

            cell137.Append(cellValue131);

            Cell cell138 = new Cell(){ CellReference = "C6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue132 = new CellValue();
            cellValue132.Text = "50";

            cell138.Append(cellValue132);

            Cell cell139 = new Cell(){ CellReference = "D6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue133 = new CellValue();
            cellValue133.Text = "51";

            cell139.Append(cellValue133);

            Cell cell140 = new Cell(){ CellReference = "E6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue134 = new CellValue();
            cellValue134.Text = "52";

            cell140.Append(cellValue134);

            Cell cell141 = new Cell(){ CellReference = "F6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue135 = new CellValue();
            cellValue135.Text = "31";

            cell141.Append(cellValue135);

            Cell cell142 = new Cell(){ CellReference = "G6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue136 = new CellValue();
            cellValue136.Text = "0";

            cell142.Append(cellValue136);

            Cell cell143 = new Cell(){ CellReference = "H6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue137 = new CellValue();
            cellValue137.Text = "15.37";

            cell143.Append(cellValue137);

            Cell cell144 = new Cell(){ CellReference = "I6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue138 = new CellValue();
            cellValue138.Text = "82";

            cell144.Append(cellValue138);

            Cell cell145 = new Cell(){ CellReference = "J6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue139 = new CellValue();
            cellValue139.Text = "1260.67";

            cell145.Append(cellValue139);

            Cell cell146 = new Cell(){ CellReference = "K6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue140 = new CellValue();
            cellValue140.Text = "1016.7725";

            cell146.Append(cellValue140);

            Cell cell147 = new Cell(){ CellReference = "L6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue141 = new CellValue();
            cellValue141.Text = "4067.09";

            cell147.Append(cellValue141);

            Cell cell148 = new Cell(){ CellReference = "M6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue142 = new CellValue();
            cellValue142.Text = "557.13";

            cell148.Append(cellValue142);

            Cell cell149 = new Cell(){ CellReference = "N6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue143 = new CellValue();
            cellValue143.Text = "469.32749999999999";

            cell149.Append(cellValue143);

            Cell cell150 = new Cell(){ CellReference = "O6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue144 = new CellValue();
            cellValue144.Text = "1877.31";

            cell150.Append(cellValue144);

            Cell cell151 = new Cell(){ CellReference = "P6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue145 = new CellValue();
            cellValue145.Text = "0.44190000000000002";

            cell151.Append(cellValue145);

            Cell cell152 = new Cell(){ CellReference = "Q6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue146 = new CellValue();
            cellValue146.Text = "0.461586";

            cell152.Append(cellValue146);

            Cell cell153 = new Cell(){ CellReference = "R6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue147 = new CellValue();
            cellValue147.Text = "1.46";

            cell153.Append(cellValue147);

            Cell cell154 = new Cell(){ CellReference = "S6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue148 = new CellValue();
            cellValue148.Text = "21.852499999999999";

            cell154.Append(cellValue148);

            Cell cell155 = new Cell(){ CellReference = "T6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue149 = new CellValue();
            cellValue149.Text = "87.41";

            cell155.Append(cellValue149);

            Cell cell156 = new Cell(){ CellReference = "U6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue150 = new CellValue();
            cellValue150.Text = "1.1999999999999999E-3";

            cell156.Append(cellValue150);

            Cell cell157 = new Cell(){ CellReference = "V6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue151 = new CellValue();
            cellValue151.Text = "2.1492000000000001E-2";

            cell157.Append(cellValue151);

            Cell cell158 = new Cell(){ CellReference = "W6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue152 = new CellValue();
            cellValue152.Text = "1.1638756400000001";

            cell158.Append(cellValue152);

            Cell cell159 = new Cell(){ CellReference = "X6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue153 = new CellValue();
            cellValue153.Text = "8.0052769999999995E-2";

            cell159.Append(cellValue153);

            Cell cell160 = new Cell(){ CellReference = "Y6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue154 = new CellValue();
            cellValue154.Text = "5.8615359500000004";

            cell160.Append(cellValue154);

            Cell cell161 = new Cell(){ CellReference = "Z6", StyleIndex = (UInt32Value)2U };
            CellValue cellValue155 = new CellValue();
            cellValue155.Text = "19.01719915";

            cell161.Append(cellValue155);

            Cell cell162 = new Cell(){ CellReference = "AA6", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue156 = new CellValue();
            cellValue156.Text = "32";

            cell162.Append(cellValue156);

            row6.Append(cell136);
            row6.Append(cell137);
            row6.Append(cell138);
            row6.Append(cell139);
            row6.Append(cell140);
            row6.Append(cell141);
            row6.Append(cell142);
            row6.Append(cell143);
            row6.Append(cell144);
            row6.Append(cell145);
            row6.Append(cell146);
            row6.Append(cell147);
            row6.Append(cell148);
            row6.Append(cell149);
            row6.Append(cell150);
            row6.Append(cell151);
            row6.Append(cell152);
            row6.Append(cell153);
            row6.Append(cell154);
            row6.Append(cell155);
            row6.Append(cell156);
            row6.Append(cell157);
            row6.Append(cell158);
            row6.Append(cell159);
            row6.Append(cell160);
            row6.Append(cell161);
            row6.Append(cell162);

            Row row7 = new Row(){ RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell163 = new Cell(){ CellReference = "A7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue157 = new CellValue();
            cellValue157.Text = "53";

            cell163.Append(cellValue157);

            Cell cell164 = new Cell(){ CellReference = "B7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue158 = new CellValue();
            cellValue158.Text = "49";

            cell164.Append(cellValue158);

            Cell cell165 = new Cell(){ CellReference = "C7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue159 = new CellValue();
            cellValue159.Text = "53";

            cell165.Append(cellValue159);

            Cell cell166 = new Cell(){ CellReference = "D7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue160 = new CellValue();
            cellValue160.Text = "54";

            cell166.Append(cellValue160);

            Cell cell167 = new Cell(){ CellReference = "E7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue161 = new CellValue();
            cellValue161.Text = "55";

            cell167.Append(cellValue161);

            Cell cell168 = new Cell(){ CellReference = "F7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue162 = new CellValue();
            cellValue162.Text = "37";

            cell168.Append(cellValue162);

            Cell cell169 = new Cell(){ CellReference = "G7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue163 = new CellValue();
            cellValue163.Text = "0";

            cell169.Append(cellValue163);

            Cell cell170 = new Cell(){ CellReference = "H7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue164 = new CellValue();
            cellValue164.Text = "12.73";

            cell170.Append(cellValue164);

            Cell cell171 = new Cell(){ CellReference = "I7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue165 = new CellValue();
            cellValue165.Text = "64";

            cell171.Append(cellValue165);

            Cell cell172 = new Cell(){ CellReference = "J7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue166 = new CellValue();
            cellValue166.Text = "814.99";

            cell172.Append(cellValue166);

            Cell cell173 = new Cell(){ CellReference = "K7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue167 = new CellValue();
            cellValue167.Text = "1182.8900000000001";

            cell173.Append(cellValue167);

            Cell cell174 = new Cell(){ CellReference = "L7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue168 = new CellValue();
            cellValue168.Text = "4731.5600000000004";

            cell174.Append(cellValue168);

            Cell cell175 = new Cell(){ CellReference = "M7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue169 = new CellValue();
            cellValue169.Text = "399.62";

            cell175.Append(cellValue169);

            Cell cell176 = new Cell(){ CellReference = "N7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue170 = new CellValue();
            cellValue170.Text = "500.58749999999998";

            cell176.Append(cellValue170);

            Cell cell177 = new Cell(){ CellReference = "O7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue171 = new CellValue();
            cellValue171.Text = "2002.35";

            cell177.Append(cellValue171);

            Cell cell178 = new Cell(){ CellReference = "P7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue172 = new CellValue();
            cellValue172.Text = "0.49030000000000001";

            cell178.Append(cellValue172);

            Cell cell179 = new Cell(){ CellReference = "Q7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue173 = new CellValue();
            cellValue173.Text = "0.42319000000000001";

            cell179.Append(cellValue173);

            Cell cell180 = new Cell(){ CellReference = "R7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue174 = new CellValue();
            cellValue174.Text = "32";

            cell180.Append(cellValue174);

            Cell cell181 = new Cell(){ CellReference = "S7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue175 = new CellValue();
            cellValue175.Text = "29.675000000000001";

            cell181.Append(cellValue175);

            Cell cell182 = new Cell(){ CellReference = "T7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue176 = new CellValue();
            cellValue176.Text = "118.7";

            cell182.Append(cellValue176);

            Cell cell183 = new Cell(){ CellReference = "U7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue177 = new CellValue();
            cellValue177.Text = "3.9300000000000002E-2";

            cell183.Append(cellValue177);

            Cell cell184 = new Cell(){ CellReference = "V7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue178 = new CellValue();
            cellValue178.Text = "2.5087000000000002E-2";

            cell184.Append(cellValue178);

            Cell cell185 = new Cell(){ CellReference = "W7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue179 = new CellValue();
            cellValue179.Text = "1.8507105500000001";

            cell185.Append(cellValue179);

            Cell cell186 = new Cell(){ CellReference = "X7", StyleIndex = (UInt32Value)2U };
            CellValue cellValue180 = new CellValue();
            cellValue180.Text = "0.2040747";

            cell186.Append(cellValue180);
            Cell cell187 = new Cell(){ CellReference = "Y7", StyleIndex = (UInt32Value)2U };
            Cell cell188 = new Cell(){ CellReference = "Z7", StyleIndex = (UInt32Value)2U };

            Cell cell189 = new Cell(){ CellReference = "AA7", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue181 = new CellValue();
            cellValue181.Text = "32";

            cell189.Append(cellValue181);

            row7.Append(cell163);
            row7.Append(cell164);
            row7.Append(cell165);
            row7.Append(cell166);
            row7.Append(cell167);
            row7.Append(cell168);
            row7.Append(cell169);
            row7.Append(cell170);
            row7.Append(cell171);
            row7.Append(cell172);
            row7.Append(cell173);
            row7.Append(cell174);
            row7.Append(cell175);
            row7.Append(cell176);
            row7.Append(cell177);
            row7.Append(cell178);
            row7.Append(cell179);
            row7.Append(cell180);
            row7.Append(cell181);
            row7.Append(cell182);
            row7.Append(cell183);
            row7.Append(cell184);
            row7.Append(cell185);
            row7.Append(cell186);
            row7.Append(cell187);
            row7.Append(cell188);
            row7.Append(cell189);

            Row row8 = new Row(){ RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell190 = new Cell(){ CellReference = "A8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue182 = new CellValue();
            cellValue182.Text = "56";

            cell190.Append(cellValue182);

            Cell cell191 = new Cell(){ CellReference = "B8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue183 = new CellValue();
            cellValue183.Text = "57";

            cell191.Append(cellValue183);

            Cell cell192 = new Cell(){ CellReference = "C8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue184 = new CellValue();
            cellValue184.Text = "58";

            cell192.Append(cellValue184);

            Cell cell193 = new Cell(){ CellReference = "D8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue185 = new CellValue();
            cellValue185.Text = "59";

            cell193.Append(cellValue185);

            Cell cell194 = new Cell(){ CellReference = "E8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue186 = new CellValue();
            cellValue186.Text = "60";

            cell194.Append(cellValue186);

            Cell cell195 = new Cell(){ CellReference = "F8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue187 = new CellValue();
            cellValue187.Text = "37";

            cell195.Append(cellValue187);

            Cell cell196 = new Cell(){ CellReference = "G8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue188 = new CellValue();
            cellValue188.Text = "0";

            cell196.Append(cellValue188);

            Cell cell197 = new Cell(){ CellReference = "H8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue189 = new CellValue();
            cellValue189.Text = "8.6199999999999992";

            cell197.Append(cellValue189);

            Cell cell198 = new Cell(){ CellReference = "I8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue190 = new CellValue();
            cellValue190.Text = "14";

            cell198.Append(cellValue190);

            Cell cell199 = new Cell(){ CellReference = "J8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue191 = new CellValue();
            cellValue191.Text = "120.66";

            cell199.Append(cellValue191);

            Cell cell200 = new Cell(){ CellReference = "K8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue192 = new CellValue();
            cellValue192.Text = "307.91199999999998";

            cell200.Append(cellValue192);

            Cell cell201 = new Cell(){ CellReference = "L8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue193 = new CellValue();
            cellValue193.Text = "1539.56";

            cell201.Append(cellValue193);

            Cell cell202 = new Cell(){ CellReference = "M8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue194 = new CellValue();
            cellValue194.Text = "87.42";

            cell202.Append(cellValue194);

            Cell cell203 = new Cell(){ CellReference = "N8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue195 = new CellValue();
            cellValue195.Text = "140.934";

            cell203.Append(cellValue195);

            Cell cell204 = new Cell(){ CellReference = "O8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue196 = new CellValue();
            cellValue196.Text = "704.67";

            cell204.Append(cellValue196);

            Cell cell205 = new Cell(){ CellReference = "P8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue197 = new CellValue();
            cellValue197.Text = "0.72450000000000003";

            cell205.Append(cellValue197);

            Cell cell206 = new Cell(){ CellReference = "Q8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue198 = new CellValue();
            cellValue198.Text = "0.45770899999999998";

            cell206.Append(cellValue198);

            Cell cell207 = new Cell(){ CellReference = "R8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue199 = new CellValue();
            cellValue199.Text = "4.0999999999999996";

            cell207.Append(cellValue199);

            Cell cell208 = new Cell(){ CellReference = "S8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue200 = new CellValue();
            cellValue200.Text = "7.94";

            cell208.Append(cellValue200);

            Cell cell209 = new Cell(){ CellReference = "T8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue201 = new CellValue();
            cellValue201.Text = "39.700000000000003";

            cell209.Append(cellValue201);

            Cell cell210 = new Cell(){ CellReference = "U8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue202 = new CellValue();
            cellValue202.Text = "3.4000000000000002E-2";

            cell210.Append(cellValue202);

            Cell cell211 = new Cell(){ CellReference = "V8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue203 = new CellValue();
            cellValue203.Text = "2.5787000000000001E-2";

            cell211.Append(cellValue203);

            Cell cell212 = new Cell(){ CellReference = "W8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue204 = new CellValue();
            cellValue204.Text = "3.3815369099999999";

            cell212.Append(cellValue204);

            Cell cell213 = new Cell(){ CellReference = "X8", StyleIndex = (UInt32Value)2U };
            CellValue cellValue205 = new CellValue();
            cellValue205.Text = "1.0893506500000001";

            cell213.Append(cellValue205);
            Cell cell214 = new Cell(){ CellReference = "Y8", StyleIndex = (UInt32Value)2U };
            Cell cell215 = new Cell(){ CellReference = "Z8", StyleIndex = (UInt32Value)2U };

            Cell cell216 = new Cell(){ CellReference = "AA8", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue206 = new CellValue();
            cellValue206.Text = "32";

            cell216.Append(cellValue206);

            row8.Append(cell190);
            row8.Append(cell191);
            row8.Append(cell192);
            row8.Append(cell193);
            row8.Append(cell194);
            row8.Append(cell195);
            row8.Append(cell196);
            row8.Append(cell197);
            row8.Append(cell198);
            row8.Append(cell199);
            row8.Append(cell200);
            row8.Append(cell201);
            row8.Append(cell202);
            row8.Append(cell203);
            row8.Append(cell204);
            row8.Append(cell205);
            row8.Append(cell206);
            row8.Append(cell207);
            row8.Append(cell208);
            row8.Append(cell209);
            row8.Append(cell210);
            row8.Append(cell211);
            row8.Append(cell212);
            row8.Append(cell213);
            row8.Append(cell214);
            row8.Append(cell215);
            row8.Append(cell216);

            Row row9 = new Row(){ RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell217 = new Cell(){ CellReference = "A9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue207 = new CellValue();
            cellValue207.Text = "56";

            cell217.Append(cellValue207);

            Cell cell218 = new Cell(){ CellReference = "B9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue208 = new CellValue();
            cellValue208.Text = "57";

            cell218.Append(cellValue208);

            Cell cell219 = new Cell(){ CellReference = "C9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue209 = new CellValue();
            cellValue209.Text = "61";

            cell219.Append(cellValue209);

            Cell cell220 = new Cell(){ CellReference = "D9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue210 = new CellValue();
            cellValue210.Text = "62";

            cell220.Append(cellValue210);

            Cell cell221 = new Cell(){ CellReference = "E9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue211 = new CellValue();
            cellValue211.Text = "63";

            cell221.Append(cellValue211);

            Cell cell222 = new Cell(){ CellReference = "F9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue212 = new CellValue();
            cellValue212.Text = "37";

            cell222.Append(cellValue212);

            Cell cell223 = new Cell(){ CellReference = "G9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue213 = new CellValue();
            cellValue213.Text = "0";

            cell223.Append(cellValue213);

            Cell cell224 = new Cell(){ CellReference = "H9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue214 = new CellValue();
            cellValue214.Text = "13.04";

            cell224.Append(cellValue214);

            Cell cell225 = new Cell(){ CellReference = "I9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue215 = new CellValue();
            cellValue215.Text = "35";

            cell225.Append(cellValue215);

            Cell cell226 = new Cell(){ CellReference = "J9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue216 = new CellValue();
            cellValue216.Text = "456.49";

            cell226.Append(cellValue216);

            Cell cell227 = new Cell(){ CellReference = "K9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue217 = new CellValue();
            cellValue217.Text = "1610.425";

            cell227.Append(cellValue217);

            Cell cell228 = new Cell(){ CellReference = "L9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue218 = new CellValue();
            cellValue218.Text = "6441.7";

            cell228.Append(cellValue218);

            Cell cell229 = new Cell(){ CellReference = "M9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue219 = new CellValue();
            cellValue219.Text = "237.24";

            cell229.Append(cellValue219);

            Cell cell230 = new Cell(){ CellReference = "N9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue220 = new CellValue();
            cellValue220.Text = "736.74";

            cell230.Append(cellValue220);

            Cell cell231 = new Cell(){ CellReference = "O9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue221 = new CellValue();
            cellValue221.Text = "2946.96";

            cell231.Append(cellValue221);

            Cell cell232 = new Cell(){ CellReference = "P9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue222 = new CellValue();
            cellValue222.Text = "0.51970000000000005";

            cell232.Append(cellValue222);

            Cell cell233 = new Cell(){ CellReference = "Q9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue223 = new CellValue();
            cellValue223.Text = "0.457482";

            cell233.Append(cellValue223);

            Cell cell234 = new Cell(){ CellReference = "R9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue224 = new CellValue();
            cellValue224.Text = "29.3";

            cell234.Append(cellValue224);

            Cell cell235 = new Cell(){ CellReference = "S9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue225 = new CellValue();
            cellValue225.Text = "45.322499999999998";

            cell235.Append(cellValue225);

            Cell cell236 = new Cell(){ CellReference = "T9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue226 = new CellValue();
            cellValue226.Text = "181.29";

            cell236.Append(cellValue226);

            Cell cell237 = new Cell(){ CellReference = "U9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue227 = new CellValue();
            cellValue227.Text = "6.4199999999999993E-2";

            cell237.Append(cellValue227);

            Cell cell238 = new Cell(){ CellReference = "V9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue228 = new CellValue();
            cellValue228.Text = "2.8143000000000001E-2";

            cell238.Append(cellValue228);

            Cell cell239 = new Cell(){ CellReference = "W9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue229 = new CellValue();
            cellValue229.Text = "1.0274903799999999";

            cell239.Append(cellValue229);

            Cell cell240 = new Cell(){ CellReference = "X9", StyleIndex = (UInt32Value)2U };
            CellValue cellValue230 = new CellValue();
            cellValue230.Text = "5.0174839999999998E-2";

            cell240.Append(cellValue230);
            Cell cell241 = new Cell(){ CellReference = "Y9", StyleIndex = (UInt32Value)2U };
            Cell cell242 = new Cell(){ CellReference = "Z9", StyleIndex = (UInt32Value)2U };

            Cell cell243 = new Cell(){ CellReference = "AA9", StyleIndex = (UInt32Value)2U, DataType = CellValues.SharedString };
            CellValue cellValue231 = new CellValue();
            cellValue231.Text = "32";

            cell243.Append(cellValue231);

            row9.Append(cell217);
            row9.Append(cell218);
            row9.Append(cell219);
            row9.Append(cell220);
            row9.Append(cell221);
            row9.Append(cell222);
            row9.Append(cell223);
            row9.Append(cell224);
            row9.Append(cell225);
            row9.Append(cell226);
            row9.Append(cell227);
            row9.Append(cell228);
            row9.Append(cell229);
            row9.Append(cell230);
            row9.Append(cell231);
            row9.Append(cell232);
            row9.Append(cell233);
            row9.Append(cell234);
            row9.Append(cell235);
            row9.Append(cell236);
            row9.Append(cell237);
            row9.Append(cell238);
            row9.Append(cell239);
            row9.Append(cell240);
            row9.Append(cell241);
            row9.Append(cell242);
            row9.Append(cell243);

            Row row10 = new Row(){ RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell244 = new Cell(){ CellReference = "E12", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue232 = new CellValue();
            cellValue232.Text = "64";

            cell244.Append(cellValue232);

            Cell cell245 = new Cell(){ CellReference = "G12", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue233 = new CellValue();
            cellValue233.Text = "26";

            cell245.Append(cellValue233);

            row10.Append(cell244);
            row10.Append(cell245);

            Row row11 = new Row(){ RowIndex = (UInt32Value)13U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell246 = new Cell(){ CellReference = "E13", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue234 = new CellValue();
            cellValue234.Text = "2";

            cell246.Append(cellValue234);

            Cell cell247 = new Cell(){ CellReference = "F13", StyleIndex = (UInt32Value)3U, DataType = CellValues.SharedString };
            CellValue cellValue235 = new CellValue();
            cellValue235.Text = "5";

            cell247.Append(cellValue235);

            Cell cell248 = new Cell(){ CellReference = "G13", DataType = CellValues.SharedString };
            CellValue cellValue236 = new CellValue();
            cellValue236.Text = "32";

            cell248.Append(cellValue236);

            Cell cell249 = new Cell(){ CellReference = "H13", DataType = CellValues.SharedString };
            CellValue cellValue237 = new CellValue();
            cellValue237.Text = "65";

            cell249.Append(cellValue237);

            Cell cell250 = new Cell(){ CellReference = "I13", DataType = CellValues.SharedString };
            CellValue cellValue238 = new CellValue();
            cellValue238.Text = "66";

            cell250.Append(cellValue238);

            row11.Append(cell246);
            row11.Append(cell247);
            row11.Append(cell248);
            row11.Append(cell249);
            row11.Append(cell250);

            Row row12 = new Row(){ RowIndex = (UInt32Value)14U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell251 = new Cell(){ CellReference = "E14", DataType = CellValues.SharedString };
            CellValue cellValue239 = new CellValue();
            cellValue239.Text = "34";

            cell251.Append(cellValue239);

            Cell cell252 = new Cell(){ CellReference = "F14", DataType = CellValues.SharedString };
            CellValue cellValue240 = new CellValue();
            cellValue240.Text = "37";

            cell252.Append(cellValue240);

            Cell cell253 = new Cell(){ CellReference = "G14", StyleIndex = (UInt32Value)4U };
            CellValue cellValue241 = new CellValue();
            cellValue241.Text = "-0.50734760999999995";

            cell253.Append(cellValue241);
            Cell cell254 = new Cell(){ CellReference = "H14", StyleIndex = (UInt32Value)4U };

            Cell cell255 = new Cell(){ CellReference = "I14", StyleIndex = (UInt32Value)4U };
            CellValue cellValue242 = new CellValue();
            cellValue242.Text = "-0.50734760999999995";

            cell255.Append(cellValue242);

            row12.Append(cell251);
            row12.Append(cell252);
            row12.Append(cell253);
            row12.Append(cell254);
            row12.Append(cell255);

            Row row13 = new Row(){ RowIndex = (UInt32Value)15U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell256 = new Cell(){ CellReference = "E15", DataType = CellValues.SharedString };
            CellValue cellValue243 = new CellValue();
            cellValue243.Text = "67";

            cell256.Append(cellValue243);

            Cell cell257 = new Cell(){ CellReference = "G15", StyleIndex = (UInt32Value)4U };
            CellValue cellValue244 = new CellValue();
            cellValue244.Text = "-0.50734760999999995";

            cell257.Append(cellValue244);
            Cell cell258 = new Cell(){ CellReference = "H15", StyleIndex = (UInt32Value)4U };

            Cell cell259 = new Cell(){ CellReference = "I15", StyleIndex = (UInt32Value)4U };
            CellValue cellValue245 = new CellValue();
            cellValue245.Text = "-0.50734760999999995";

            cell259.Append(cellValue245);

            row13.Append(cell256);
            row13.Append(cell257);
            row13.Append(cell258);
            row13.Append(cell259);

            Row row14 = new Row(){ RowIndex = (UInt32Value)16U, Spans = new ListValue<StringValue>() { InnerText = "1:27" } };

            Cell cell260 = new Cell(){ CellReference = "E16", DataType = CellValues.SharedString };
            CellValue cellValue246 = new CellValue();
            cellValue246.Text = "53";

            cell260.Append(cellValue246);

            Cell cell261 = new Cell(){ CellReference = "F16", DataType = CellValues.SharedString };
            CellValue cellValue247 = new CellValue();
            cellValue247.Text = "37";

            cell261.Append(cellValue247);
            Cell cell262 = new Cell(){ CellReference = "G16", StyleIndex = (UInt32Value)4U };
            Cell cell263 = new Cell(){ CellReference = "H16", StyleIndex = (UInt32Value)4U };
            Cell cell264 = new Cell(){ CellReference = "I16", StyleIndex = (UInt32Value)4U };

            row14.Append(cell260);
            row14.Append(cell261);
            row14.Append(cell262);
            row14.Append(cell263);
            row14.Append(cell264);

            Row row15 = new Row(){ RowIndex = (UInt32Value)17U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell265 = new Cell(){ CellReference = "E17", DataType = CellValues.SharedString };
            CellValue cellValue248 = new CellValue();
            cellValue248.Text = "68";

            cell265.Append(cellValue248);
            Cell cell266 = new Cell(){ CellReference = "G17", StyleIndex = (UInt32Value)4U };
            Cell cell267 = new Cell(){ CellReference = "H17", StyleIndex = (UInt32Value)4U };
            Cell cell268 = new Cell(){ CellReference = "I17", StyleIndex = (UInt32Value)4U };

            row15.Append(cell265);
            row15.Append(cell266);
            row15.Append(cell267);
            row15.Append(cell268);

            Row row16 = new Row(){ RowIndex = (UInt32Value)18U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell269 = new Cell(){ CellReference = "E18", DataType = CellValues.SharedString };
            CellValue cellValue249 = new CellValue();
            cellValue249.Text = "58";

            cell269.Append(cellValue249);

            Cell cell270 = new Cell(){ CellReference = "F18", DataType = CellValues.SharedString };
            CellValue cellValue250 = new CellValue();
            cellValue250.Text = "37";

            cell270.Append(cellValue250);
            Cell cell271 = new Cell(){ CellReference = "G18", StyleIndex = (UInt32Value)4U };
            Cell cell272 = new Cell(){ CellReference = "H18", StyleIndex = (UInt32Value)4U };
            Cell cell273 = new Cell(){ CellReference = "I18", StyleIndex = (UInt32Value)4U };

            row16.Append(cell269);
            row16.Append(cell270);
            row16.Append(cell271);
            row16.Append(cell272);
            row16.Append(cell273);

            Row row17 = new Row(){ RowIndex = (UInt32Value)19U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell274 = new Cell(){ CellReference = "E19", DataType = CellValues.SharedString };
            CellValue cellValue251 = new CellValue();
            cellValue251.Text = "69";

            cell274.Append(cellValue251);
            Cell cell275 = new Cell(){ CellReference = "G19", StyleIndex = (UInt32Value)4U };
            Cell cell276 = new Cell(){ CellReference = "H19", StyleIndex = (UInt32Value)4U };
            Cell cell277 = new Cell(){ CellReference = "I19", StyleIndex = (UInt32Value)4U };

            row17.Append(cell274);
            row17.Append(cell275);
            row17.Append(cell276);
            row17.Append(cell277);

            Row row18 = new Row(){ RowIndex = (UInt32Value)20U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell278 = new Cell(){ CellReference = "E20", DataType = CellValues.SharedString };
            CellValue cellValue252 = new CellValue();
            cellValue252.Text = "61";

            cell278.Append(cellValue252);

            Cell cell279 = new Cell(){ CellReference = "F20", DataType = CellValues.SharedString };
            CellValue cellValue253 = new CellValue();
            cellValue253.Text = "37";

            cell279.Append(cellValue253);
            Cell cell280 = new Cell(){ CellReference = "G20", StyleIndex = (UInt32Value)4U };
            Cell cell281 = new Cell(){ CellReference = "H20", StyleIndex = (UInt32Value)4U };
            Cell cell282 = new Cell(){ CellReference = "I20", StyleIndex = (UInt32Value)4U };

            row18.Append(cell278);
            row18.Append(cell279);
            row18.Append(cell280);
            row18.Append(cell281);
            row18.Append(cell282);

            Row row19 = new Row(){ RowIndex = (UInt32Value)21U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell283 = new Cell(){ CellReference = "E21", DataType = CellValues.SharedString };
            CellValue cellValue254 = new CellValue();
            cellValue254.Text = "70";

            cell283.Append(cellValue254);
            Cell cell284 = new Cell(){ CellReference = "G21", StyleIndex = (UInt32Value)4U };
            Cell cell285 = new Cell(){ CellReference = "H21", StyleIndex = (UInt32Value)4U };
            Cell cell286 = new Cell(){ CellReference = "I21", StyleIndex = (UInt32Value)4U };

            row19.Append(cell283);
            row19.Append(cell284);
            row19.Append(cell285);
            row19.Append(cell286);

            Row row20 = new Row(){ RowIndex = (UInt32Value)22U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell287 = new Cell(){ CellReference = "E22", DataType = CellValues.SharedString };
            CellValue cellValue255 = new CellValue();
            cellValue255.Text = "40";

            cell287.Append(cellValue255);

            Cell cell288 = new Cell(){ CellReference = "F22", DataType = CellValues.SharedString };
            CellValue cellValue256 = new CellValue();
            cellValue256.Text = "31";

            cell288.Append(cellValue256);
            Cell cell289 = new Cell(){ CellReference = "G22", StyleIndex = (UInt32Value)4U };
            Cell cell290 = new Cell(){ CellReference = "H22", StyleIndex = (UInt32Value)4U };
            Cell cell291 = new Cell(){ CellReference = "I22", StyleIndex = (UInt32Value)4U };

            row20.Append(cell287);
            row20.Append(cell288);
            row20.Append(cell289);
            row20.Append(cell290);
            row20.Append(cell291);

            Row row21 = new Row(){ RowIndex = (UInt32Value)23U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell292 = new Cell(){ CellReference = "E23", DataType = CellValues.SharedString };
            CellValue cellValue257 = new CellValue();
            cellValue257.Text = "71";

            cell292.Append(cellValue257);
            Cell cell293 = new Cell(){ CellReference = "G23", StyleIndex = (UInt32Value)4U };
            Cell cell294 = new Cell(){ CellReference = "H23", StyleIndex = (UInt32Value)4U };
            Cell cell295 = new Cell(){ CellReference = "I23", StyleIndex = (UInt32Value)4U };

            row21.Append(cell292);
            row21.Append(cell293);
            row21.Append(cell294);
            row21.Append(cell295);

            Row row22 = new Row(){ RowIndex = (UInt32Value)24U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell296 = new Cell(){ CellReference = "E24", DataType = CellValues.SharedString };
            CellValue cellValue258 = new CellValue();
            cellValue258.Text = "50";

            cell296.Append(cellValue258);

            Cell cell297 = new Cell(){ CellReference = "F24", DataType = CellValues.SharedString };
            CellValue cellValue259 = new CellValue();
            cellValue259.Text = "31";

            cell297.Append(cellValue259);

            Cell cell298 = new Cell(){ CellReference = "G24", StyleIndex = (UInt32Value)4U };
            CellValue cellValue260 = new CellValue();
            cellValue260.Text = "5.8615359500000004";

            cell298.Append(cellValue260);
            Cell cell299 = new Cell(){ CellReference = "H24", StyleIndex = (UInt32Value)4U };

            Cell cell300 = new Cell(){ CellReference = "I24", StyleIndex = (UInt32Value)4U };
            CellValue cellValue261 = new CellValue();
            cellValue261.Text = "5.8615359500000004";

            cell300.Append(cellValue261);

            row22.Append(cell296);
            row22.Append(cell297);
            row22.Append(cell298);
            row22.Append(cell299);
            row22.Append(cell300);

            Row row23 = new Row(){ RowIndex = (UInt32Value)25U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell301 = new Cell(){ CellReference = "E25", DataType = CellValues.SharedString };
            CellValue cellValue262 = new CellValue();
            cellValue262.Text = "72";

            cell301.Append(cellValue262);

            Cell cell302 = new Cell(){ CellReference = "G25", StyleIndex = (UInt32Value)4U };
            CellValue cellValue263 = new CellValue();
            cellValue263.Text = "5.8615359500000004";

            cell302.Append(cellValue263);
            Cell cell303 = new Cell(){ CellReference = "H25", StyleIndex = (UInt32Value)4U };

            Cell cell304 = new Cell(){ CellReference = "I25", StyleIndex = (UInt32Value)4U };
            CellValue cellValue264 = new CellValue();
            cellValue264.Text = "5.8615359500000004";

            cell304.Append(cellValue264);

            row23.Append(cell301);
            row23.Append(cell302);
            row23.Append(cell303);
            row23.Append(cell304);

            Row row24 = new Row(){ RowIndex = (UInt32Value)26U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell305 = new Cell(){ CellReference = "E26", DataType = CellValues.SharedString };
            CellValue cellValue265 = new CellValue();
            cellValue265.Text = "45";

            cell305.Append(cellValue265);

            Cell cell306 = new Cell(){ CellReference = "F26", DataType = CellValues.SharedString };
            CellValue cellValue266 = new CellValue();
            cellValue266.Text = "37";

            cell306.Append(cellValue266);
            Cell cell307 = new Cell(){ CellReference = "G26", StyleIndex = (UInt32Value)4U };
            Cell cell308 = new Cell(){ CellReference = "H26", StyleIndex = (UInt32Value)4U };
            Cell cell309 = new Cell(){ CellReference = "I26", StyleIndex = (UInt32Value)4U };

            row24.Append(cell305);
            row24.Append(cell306);
            row24.Append(cell307);
            row24.Append(cell308);
            row24.Append(cell309);

            Row row25 = new Row(){ RowIndex = (UInt32Value)27U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell310 = new Cell(){ CellReference = "E27", DataType = CellValues.SharedString };
            CellValue cellValue267 = new CellValue();
            cellValue267.Text = "73";

            cell310.Append(cellValue267);
            Cell cell311 = new Cell(){ CellReference = "G27", StyleIndex = (UInt32Value)4U };
            Cell cell312 = new Cell(){ CellReference = "H27", StyleIndex = (UInt32Value)4U };
            Cell cell313 = new Cell(){ CellReference = "I27", StyleIndex = (UInt32Value)4U };

            row25.Append(cell310);
            row25.Append(cell311);
            row25.Append(cell312);
            row25.Append(cell313);

            Row row26 = new Row(){ RowIndex = (UInt32Value)28U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell314 = new Cell(){ CellReference = "E28", DataType = CellValues.SharedString };
            CellValue cellValue268 = new CellValue();
            cellValue268.Text = "27";

            cell314.Append(cellValue268);

            Cell cell315 = new Cell(){ CellReference = "F28", DataType = CellValues.SharedString };
            CellValue cellValue269 = new CellValue();
            cellValue269.Text = "31";

            cell315.Append(cellValue269);
            Cell cell316 = new Cell(){ CellReference = "G28", StyleIndex = (UInt32Value)4U };
            Cell cell317 = new Cell(){ CellReference = "H28", StyleIndex = (UInt32Value)4U };
            Cell cell318 = new Cell(){ CellReference = "I28", StyleIndex = (UInt32Value)4U };

            row26.Append(cell314);
            row26.Append(cell315);
            row26.Append(cell316);
            row26.Append(cell317);
            row26.Append(cell318);

            Row row27 = new Row(){ RowIndex = (UInt32Value)29U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell319 = new Cell(){ CellReference = "E29", DataType = CellValues.SharedString };
            CellValue cellValue270 = new CellValue();
            cellValue270.Text = "74";

            cell319.Append(cellValue270);
            Cell cell320 = new Cell(){ CellReference = "G29", StyleIndex = (UInt32Value)4U };
            Cell cell321 = new Cell(){ CellReference = "H29", StyleIndex = (UInt32Value)4U };
            Cell cell322 = new Cell(){ CellReference = "I29", StyleIndex = (UInt32Value)4U };

            row27.Append(cell319);
            row27.Append(cell320);
            row27.Append(cell321);
            row27.Append(cell322);

            Row row28 = new Row(){ RowIndex = (UInt32Value)30U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell323 = new Cell(){ CellReference = "E30", DataType = CellValues.SharedString };
            CellValue cellValue271 = new CellValue();
            cellValue271.Text = "65";

            cell323.Append(cellValue271);

            Cell cell324 = new Cell(){ CellReference = "F30", DataType = CellValues.SharedString };
            CellValue cellValue272 = new CellValue();
            cellValue272.Text = "65";

            cell324.Append(cellValue272);
            Cell cell325 = new Cell(){ CellReference = "G30", StyleIndex = (UInt32Value)4U };
            Cell cell326 = new Cell(){ CellReference = "H30", StyleIndex = (UInt32Value)4U };
            Cell cell327 = new Cell(){ CellReference = "I30", StyleIndex = (UInt32Value)4U };

            row28.Append(cell323);
            row28.Append(cell324);
            row28.Append(cell325);
            row28.Append(cell326);
            row28.Append(cell327);

            Row row29 = new Row(){ RowIndex = (UInt32Value)31U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell328 = new Cell(){ CellReference = "E31", DataType = CellValues.SharedString };
            CellValue cellValue273 = new CellValue();
            cellValue273.Text = "75";

            cell328.Append(cellValue273);
            Cell cell329 = new Cell(){ CellReference = "G31", StyleIndex = (UInt32Value)4U };
            Cell cell330 = new Cell(){ CellReference = "H31", StyleIndex = (UInt32Value)4U };
            Cell cell331 = new Cell(){ CellReference = "I31", StyleIndex = (UInt32Value)4U };

            row29.Append(cell328);
            row29.Append(cell329);
            row29.Append(cell330);
            row29.Append(cell331);

            Row row30 = new Row(){ RowIndex = (UInt32Value)32U, Spans = new ListValue<StringValue>() { InnerText = "5:9" } };

            Cell cell332 = new Cell(){ CellReference = "E32", DataType = CellValues.SharedString };
            CellValue cellValue274 = new CellValue();
            cellValue274.Text = "66";

            cell332.Append(cellValue274);

            Cell cell333 = new Cell(){ CellReference = "G32", StyleIndex = (UInt32Value)4U };
            CellValue cellValue275 = new CellValue();
            cellValue275.Text = "5.3541883400000003";

            cell333.Append(cellValue275);
            Cell cell334 = new Cell(){ CellReference = "H32", StyleIndex = (UInt32Value)4U };

            Cell cell335 = new Cell(){ CellReference = "I32", StyleIndex = (UInt32Value)4U };
            CellValue cellValue276 = new CellValue();
            cellValue276.Text = "5.3541883400000003";

            cell335.Append(cellValue276);

            row30.Append(cell332);
            row30.Append(cell333);
            row30.Append(cell334);
            row30.Append(cell335);

            sheetData1.Append(row1);
            sheetData1.Append(row2);
            sheetData1.Append(row3);
            sheetData1.Append(row4);
            sheetData1.Append(row5);
            sheetData1.Append(row6);
            sheetData1.Append(row7);
            sheetData1.Append(row8);
            sheetData1.Append(row9);
            sheetData1.Append(row10);
            sheetData1.Append(row11);
            sheetData1.Append(row12);
            sheetData1.Append(row13);
            sheetData1.Append(row14);
            sheetData1.Append(row15);
            sheetData1.Append(row16);
            sheetData1.Append(row17);
            sheetData1.Append(row18);
            sheetData1.Append(row19);
            sheetData1.Append(row20);
            sheetData1.Append(row21);
            sheetData1.Append(row22);
            sheetData1.Append(row23);
            sheetData1.Append(row24);
            sheetData1.Append(row25);
            sheetData1.Append(row26);
            sheetData1.Append(row27);
            sheetData1.Append(row28);
            sheetData1.Append(row29);
            sheetData1.Append(row30);
            PageMargins pageMargins1 = new PageMargins(){ Left = 0.75D, Right = 0.75D, Top = 1D, Bottom = 1D, Header = 0.5D, Footer = 0.5D };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(pageMargins1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of pivotTablePart1.
        private void GeneratePivotTablePart1Content(PivotTablePart pivotTablePart1)
        {
            PivotTableDefinition pivotTableDefinition1 = new PivotTableDefinition(){ Name = "PivotTable2", CacheId = (UInt32Value)49U, ApplyNumberFormats = false, ApplyBorderFormats = false, ApplyFontFormats = false, ApplyPatternFormats = false, ApplyAlignmentFormats = false, ApplyWidthHeightFormats = true, DataCaption = "Values", UpdatedVersion = 7, MinRefreshableVersion = 3, UseAutoFormatting = true, ItemPrintTitles = true, CreatedVersion = 7, Indent = (UInt32Value)0U, Compact = false, CompactData = false, MultipleFieldFilters = false, MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "xr" }  };
            pivotTableDefinition1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            pivotTableDefinition1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");
            pivotTableDefinition1.SetAttribute(new OpenXmlAttribute("xr", "uid", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision", "{94B96C82-65DB-497E-B0B0-EC094C359F0E}"));
            Location location1 = new Location(){ Reference = "E12:I32", FirstHeaderRow = (UInt32Value)1U, FirstDataRow = (UInt32Value)2U, FirstDataColumn = (UInt32Value)2U };

            PivotFields pivotFields1 = new PivotFields(){ Count = (UInt32Value)27U };
            PivotField pivotField1 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField2 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };

            PivotField pivotField3 = new PivotField(){ Axis = PivotTableAxisValues.AxisRow, Compact = false, Outline = false, ShowAll = false };

            Items items1 = new Items(){ Count = (UInt32Value)10U };
            Item item1 = new Item(){ Index = (UInt32Value)1U };
            Item item2 = new Item(){ Index = (UInt32Value)5U };
            Item item3 = new Item(){ Index = (UInt32Value)6U };
            Item item4 = new Item(){ Index = (UInt32Value)7U };
            Item item5 = new Item(){ Index = (UInt32Value)2U };
            Item item6 = new Item(){ Index = (UInt32Value)4U };
            Item item7 = new Item(){ Index = (UInt32Value)3U };
            Item item8 = new Item(){ Index = (UInt32Value)0U };
            Item item9 = new Item(){ Index = (UInt32Value)8U };
            Item item10 = new Item(){ ItemType = ItemValues.Default };

            items1.Append(item1);
            items1.Append(item2);
            items1.Append(item3);
            items1.Append(item4);
            items1.Append(item5);
            items1.Append(item6);
            items1.Append(item7);
            items1.Append(item8);
            items1.Append(item9);
            items1.Append(item10);

            pivotField3.Append(items1);
            PivotField pivotField4 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField5 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };

            PivotField pivotField6 = new PivotField(){ Axis = PivotTableAxisValues.AxisRow, Compact = false, Outline = false, ShowAll = false };

            Items items2 = new Items(){ Count = (UInt32Value)4U };
            Item item11 = new Item(){ Index = (UInt32Value)0U };
            Item item12 = new Item(){ Index = (UInt32Value)1U };
            Item item13 = new Item(){ Index = (UInt32Value)2U };
            Item item14 = new Item(){ ItemType = ItemValues.Default };

            items2.Append(item11);
            items2.Append(item12);
            items2.Append(item13);
            items2.Append(item14);

            pivotField6.Append(items2);
            PivotField pivotField7 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField8 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField9 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField10 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField11 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField12 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField13 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField14 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField15 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField16 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField17 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField18 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField19 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField20 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField21 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField22 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField23 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField24 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField25 = new PivotField(){ DataField = true, Compact = false, Outline = false, ShowAll = false };
            PivotField pivotField26 = new PivotField(){ Compact = false, Outline = false, ShowAll = false };

            PivotField pivotField27 = new PivotField(){ Axis = PivotTableAxisValues.AxisColumn, Compact = false, Outline = false, ShowAll = false };

            Items items3 = new Items(){ Count = (UInt32Value)3U };
            Item item15 = new Item(){ Index = (UInt32Value)0U };
            Item item16 = new Item(){ Index = (UInt32Value)1U };
            Item item17 = new Item(){ ItemType = ItemValues.Default };

            items3.Append(item15);
            items3.Append(item16);
            items3.Append(item17);

            pivotField27.Append(items3);

            pivotFields1.Append(pivotField1);
            pivotFields1.Append(pivotField2);
            pivotFields1.Append(pivotField3);
            pivotFields1.Append(pivotField4);
            pivotFields1.Append(pivotField5);
            pivotFields1.Append(pivotField6);
            pivotFields1.Append(pivotField7);
            pivotFields1.Append(pivotField8);
            pivotFields1.Append(pivotField9);
            pivotFields1.Append(pivotField10);
            pivotFields1.Append(pivotField11);
            pivotFields1.Append(pivotField12);
            pivotFields1.Append(pivotField13);
            pivotFields1.Append(pivotField14);
            pivotFields1.Append(pivotField15);
            pivotFields1.Append(pivotField16);
            pivotFields1.Append(pivotField17);
            pivotFields1.Append(pivotField18);
            pivotFields1.Append(pivotField19);
            pivotFields1.Append(pivotField20);
            pivotFields1.Append(pivotField21);
            pivotFields1.Append(pivotField22);
            pivotFields1.Append(pivotField23);
            pivotFields1.Append(pivotField24);
            pivotFields1.Append(pivotField25);
            pivotFields1.Append(pivotField26);
            pivotFields1.Append(pivotField27);

            RowFields rowFields1 = new RowFields(){ Count = (UInt32Value)2U };
            Field field1 = new Field(){ Index = 2 };
            Field field2 = new Field(){ Index = 5 };

            rowFields1.Append(field1);
            rowFields1.Append(field2);

            RowItems rowItems1 = new RowItems(){ Count = (UInt32Value)19U };

            RowItem rowItem1 = new RowItem();
            MemberPropertyIndex memberPropertyIndex1 = new MemberPropertyIndex();
            MemberPropertyIndex memberPropertyIndex2 = new MemberPropertyIndex(){ Val = 1 };

            rowItem1.Append(memberPropertyIndex1);
            rowItem1.Append(memberPropertyIndex2);

            RowItem rowItem2 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex3 = new MemberPropertyIndex();

            rowItem2.Append(memberPropertyIndex3);

            RowItem rowItem3 = new RowItem();
            MemberPropertyIndex memberPropertyIndex4 = new MemberPropertyIndex(){ Val = 1 };
            MemberPropertyIndex memberPropertyIndex5 = new MemberPropertyIndex(){ Val = 1 };

            rowItem3.Append(memberPropertyIndex4);
            rowItem3.Append(memberPropertyIndex5);

            RowItem rowItem4 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex6 = new MemberPropertyIndex(){ Val = 1 };

            rowItem4.Append(memberPropertyIndex6);

            RowItem rowItem5 = new RowItem();
            MemberPropertyIndex memberPropertyIndex7 = new MemberPropertyIndex(){ Val = 2 };
            MemberPropertyIndex memberPropertyIndex8 = new MemberPropertyIndex(){ Val = 1 };

            rowItem5.Append(memberPropertyIndex7);
            rowItem5.Append(memberPropertyIndex8);

            RowItem rowItem6 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex9 = new MemberPropertyIndex(){ Val = 2 };

            rowItem6.Append(memberPropertyIndex9);

            RowItem rowItem7 = new RowItem();
            MemberPropertyIndex memberPropertyIndex10 = new MemberPropertyIndex(){ Val = 3 };
            MemberPropertyIndex memberPropertyIndex11 = new MemberPropertyIndex(){ Val = 1 };

            rowItem7.Append(memberPropertyIndex10);
            rowItem7.Append(memberPropertyIndex11);

            RowItem rowItem8 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex12 = new MemberPropertyIndex(){ Val = 3 };

            rowItem8.Append(memberPropertyIndex12);

            RowItem rowItem9 = new RowItem();
            MemberPropertyIndex memberPropertyIndex13 = new MemberPropertyIndex(){ Val = 4 };
            MemberPropertyIndex memberPropertyIndex14 = new MemberPropertyIndex();

            rowItem9.Append(memberPropertyIndex13);
            rowItem9.Append(memberPropertyIndex14);

            RowItem rowItem10 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex15 = new MemberPropertyIndex(){ Val = 4 };

            rowItem10.Append(memberPropertyIndex15);

            RowItem rowItem11 = new RowItem();
            MemberPropertyIndex memberPropertyIndex16 = new MemberPropertyIndex(){ Val = 5 };
            MemberPropertyIndex memberPropertyIndex17 = new MemberPropertyIndex();

            rowItem11.Append(memberPropertyIndex16);
            rowItem11.Append(memberPropertyIndex17);

            RowItem rowItem12 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex18 = new MemberPropertyIndex(){ Val = 5 };

            rowItem12.Append(memberPropertyIndex18);

            RowItem rowItem13 = new RowItem();
            MemberPropertyIndex memberPropertyIndex19 = new MemberPropertyIndex(){ Val = 6 };
            MemberPropertyIndex memberPropertyIndex20 = new MemberPropertyIndex(){ Val = 1 };

            rowItem13.Append(memberPropertyIndex19);
            rowItem13.Append(memberPropertyIndex20);

            RowItem rowItem14 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex21 = new MemberPropertyIndex(){ Val = 6 };

            rowItem14.Append(memberPropertyIndex21);

            RowItem rowItem15 = new RowItem();
            MemberPropertyIndex memberPropertyIndex22 = new MemberPropertyIndex(){ Val = 7 };
            MemberPropertyIndex memberPropertyIndex23 = new MemberPropertyIndex();

            rowItem15.Append(memberPropertyIndex22);
            rowItem15.Append(memberPropertyIndex23);

            RowItem rowItem16 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex24 = new MemberPropertyIndex(){ Val = 7 };

            rowItem16.Append(memberPropertyIndex24);

            RowItem rowItem17 = new RowItem();
            MemberPropertyIndex memberPropertyIndex25 = new MemberPropertyIndex(){ Val = 8 };
            MemberPropertyIndex memberPropertyIndex26 = new MemberPropertyIndex(){ Val = 2 };

            rowItem17.Append(memberPropertyIndex25);
            rowItem17.Append(memberPropertyIndex26);

            RowItem rowItem18 = new RowItem(){ ItemType = ItemValues.Default };
            MemberPropertyIndex memberPropertyIndex27 = new MemberPropertyIndex(){ Val = 8 };

            rowItem18.Append(memberPropertyIndex27);

            RowItem rowItem19 = new RowItem(){ ItemType = ItemValues.Grand };
            MemberPropertyIndex memberPropertyIndex28 = new MemberPropertyIndex();

            rowItem19.Append(memberPropertyIndex28);

            rowItems1.Append(rowItem1);
            rowItems1.Append(rowItem2);
            rowItems1.Append(rowItem3);
            rowItems1.Append(rowItem4);
            rowItems1.Append(rowItem5);
            rowItems1.Append(rowItem6);
            rowItems1.Append(rowItem7);
            rowItems1.Append(rowItem8);
            rowItems1.Append(rowItem9);
            rowItems1.Append(rowItem10);
            rowItems1.Append(rowItem11);
            rowItems1.Append(rowItem12);
            rowItems1.Append(rowItem13);
            rowItems1.Append(rowItem14);
            rowItems1.Append(rowItem15);
            rowItems1.Append(rowItem16);
            rowItems1.Append(rowItem17);
            rowItems1.Append(rowItem18);
            rowItems1.Append(rowItem19);

            ColumnFields columnFields1 = new ColumnFields(){ Count = (UInt32Value)1U };
            Field field3 = new Field(){ Index = 26 };

            columnFields1.Append(field3);

            ColumnItems columnItems1 = new ColumnItems(){ Count = (UInt32Value)3U };

            RowItem rowItem20 = new RowItem();
            MemberPropertyIndex memberPropertyIndex29 = new MemberPropertyIndex();

            rowItem20.Append(memberPropertyIndex29);

            RowItem rowItem21 = new RowItem();
            MemberPropertyIndex memberPropertyIndex30 = new MemberPropertyIndex(){ Val = 1 };

            rowItem21.Append(memberPropertyIndex30);

            RowItem rowItem22 = new RowItem(){ ItemType = ItemValues.Grand };
            MemberPropertyIndex memberPropertyIndex31 = new MemberPropertyIndex();

            rowItem22.Append(memberPropertyIndex31);

            columnItems1.Append(rowItem20);
            columnItems1.Append(rowItem21);
            columnItems1.Append(rowItem22);

            DataFields dataFields1 = new DataFields(){ Count = (UInt32Value)1U };
            DataField dataField1 = new DataField(){ Name = "Sum of 比上周", Field = (UInt32Value)24U, BaseField = 0, BaseItem = (UInt32Value)0U };

            dataFields1.Append(dataField1);
            PivotTableStyle pivotTableStyle1 = new PivotTableStyle(){ Name = "PivotStyleLight16", ShowRowHeaders = true, ShowColumnHeaders = true, ShowRowStripes = false, ShowColumnStripes = false, ShowLastColumn = true };

            PivotTableDefinitionExtensionList pivotTableDefinitionExtensionList1 = new PivotTableDefinitionExtensionList();

            PivotTableDefinitionExtension pivotTableDefinitionExtension1 = new PivotTableDefinitionExtension(){ Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" };
            pivotTableDefinitionExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");

            X14.PivotTableDefinition pivotTableDefinition2 = new X14.PivotTableDefinition(){ HideValuesRow = true };
            pivotTableDefinition2.AddNamespaceDeclaration("xm", "http://schemas.microsoft.com/office/excel/2006/main");

            pivotTableDefinitionExtension1.Append(pivotTableDefinition2);

            PivotTableDefinitionExtension pivotTableDefinitionExtension2 = new PivotTableDefinitionExtension(){ Uri = "{747A6164-185A-40DC-8AA5-F01512510D54}" };
            pivotTableDefinitionExtension2.AddNamespaceDeclaration("xpdl", "http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout");
            OpenXmlUnknownElement openXmlUnknownElement3 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<xpdl:pivotTableDefinition16 xmlns:xpdl=\"http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout\" />");

            pivotTableDefinitionExtension2.Append(openXmlUnknownElement3);

            pivotTableDefinitionExtensionList1.Append(pivotTableDefinitionExtension1);
            pivotTableDefinitionExtensionList1.Append(pivotTableDefinitionExtension2);

            pivotTableDefinition1.Append(location1);
            pivotTableDefinition1.Append(pivotFields1);
            pivotTableDefinition1.Append(rowFields1);
            pivotTableDefinition1.Append(rowItems1);
            pivotTableDefinition1.Append(columnFields1);
            pivotTableDefinition1.Append(columnItems1);
            pivotTableDefinition1.Append(dataFields1);
            pivotTableDefinition1.Append(pivotTableStyle1);
            pivotTableDefinition1.Append(pivotTableDefinitionExtensionList1);

            pivotTablePart1.PivotTableDefinition = pivotTableDefinition1;
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable(){ Count = (UInt32Value)118U, UniqueCount = (UInt32Value)76U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "城市";

            sharedStringItem1.Append(text1);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "负责人";

            sharedStringItem2.Append(text2);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "物理店";

            sharedStringItem3.Append(text3);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text4 = new Text();
            text4.Text = "门店ID";

            sharedStringItem4.Append(text4);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "门店";

            sharedStringItem5.Append(text5);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "平台";

            sharedStringItem6.Append(text6);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "三方配送";

            sharedStringItem7.Append(text7);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "单价";

            sharedStringItem8.Append(text8);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "订单";

            sharedStringItem9.Append(text9);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "收入";

            sharedStringItem10.Append(text10);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "平均收入";

            sharedStringItem11.Append(text11);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "总收入";

            sharedStringItem12.Append(text12);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "成本";

            sharedStringItem13.Append(text13);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "平均成本";

            sharedStringItem14.Append(text14);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "总成本";

            sharedStringItem15.Append(text15);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "成本比例";

            sharedStringItem16.Append(text16);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "总成本比例";

            sharedStringItem17.Append(text17);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "推广";

            sharedStringItem18.Append(text18);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "平均推广";

            sharedStringItem19.Append(text19);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "总推广";

            sharedStringItem20.Append(text20);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "推广比例";

            sharedStringItem21.Append(text21);

            SharedStringItem sharedStringItem22 = new SharedStringItem();
            Text text22 = new Text();
            text22.Text = "总推广比例";

            sharedStringItem22.Append(text22);

            SharedStringItem sharedStringItem23 = new SharedStringItem();
            Text text23 = new Text();
            text23.Text = "比30天";

            sharedStringItem23.Append(text23);

            SharedStringItem sharedStringItem24 = new SharedStringItem();
            Text text24 = new Text();
            text24.Text = "比上天";

            sharedStringItem24.Append(text24);

            SharedStringItem sharedStringItem25 = new SharedStringItem();
            Text text25 = new Text();
            text25.Text = "比上周";

            sharedStringItem25.Append(text25);

            SharedStringItem sharedStringItem26 = new SharedStringItem();
            Text text26 = new Text();
            text26.Text = "比上周(3)";

            sharedStringItem26.Append(text26);

            SharedStringItem sharedStringItem27 = new SharedStringItem();
            Text text27 = new Text();
            text27.Text = "日期";

            sharedStringItem27.Append(text27);

            SharedStringItem sharedStringItem28 = new SharedStringItem();
            Text text28 = new Text();
            text28.Text = "赣州";

            sharedStringItem28.Append(text28);

            SharedStringItem sharedStringItem29 = new SharedStringItem();
            Text text29 = new Text();
            text29.Text = "小伍";

            sharedStringItem29.Append(text29);

            SharedStringItem sharedStringItem30 = new SharedStringItem();
            Text text30 = new Text();
            text30.Text = "11027801";

            sharedStringItem30.Append(text30);

            SharedStringItem sharedStringItem31 = new SharedStringItem();
            Text text31 = new Text();
            text31.Text = "金牌手抓饼•奶茶•小吃(赣州店)";

            sharedStringItem31.Append(text31);

            SharedStringItem sharedStringItem32 = new SharedStringItem();
            Text text32 = new Text();
            text32.Text = "美团";

            sharedStringItem32.Append(text32);

            SharedStringItem sharedStringItem33 = new SharedStringItem();
            Text text33 = new Text();
            text33.Text = "20210302";

            sharedStringItem33.Append(text33);

            SharedStringItem sharedStringItem34 = new SharedStringItem();
            Text text34 = new Text();
            text34.Text = "云南";

            sharedStringItem34.Append(text34);

            SharedStringItem sharedStringItem35 = new SharedStringItem();
            Text text35 = new Text();
            text35.Text = "云南旗舰店";

            sharedStringItem35.Append(text35);

            SharedStringItem sharedStringItem36 = new SharedStringItem();
            Text text36 = new Text();
            text36.Text = "2072014729";

            sharedStringItem36.Append(text36);

            SharedStringItem sharedStringItem37 = new SharedStringItem();
            Text text37 = new Text();
            text37.Text = "苏姐牛奶甜品世家(昭通店)";

            sharedStringItem37.Append(text37);

            SharedStringItem sharedStringItem38 = new SharedStringItem();
            Text text38 = new Text();
            text38.Text = "饿了么";

            sharedStringItem38.Append(text38);

            SharedStringItem sharedStringItem39 = new SharedStringItem();
            Text text39 = new Text();
            text39.Text = "武汉";

            sharedStringItem39.Append(text39);

            SharedStringItem sharedStringItem40 = new SharedStringItem();
            Text text40 = new Text();
            text40.Text = "邓信庭";

            sharedStringItem40.Append(text40);

            SharedStringItem sharedStringItem41 = new SharedStringItem();
            Text text41 = new Text();
            text41.Text = "江汉";

            sharedStringItem41.Append(text41);

            SharedStringItem sharedStringItem42 = new SharedStringItem();
            Text text42 = new Text();
            text42.Text = "11057742";

            sharedStringItem42.Append(text42);

            SharedStringItem sharedStringItem43 = new SharedStringItem();
            Text text43 = new Text();
            text43.Text = "古御贡茶•手抓饼•小吃（江汉店）";

            sharedStringItem43.Append(text43);

            SharedStringItem sharedStringItem44 = new SharedStringItem();
            Text text44 = new Text();
            text44.Text = "广州";

            sharedStringItem44.Append(text44);

            SharedStringItem sharedStringItem45 = new SharedStringItem();
            Text text45 = new Text();
            text45.Text = "郑秀娟";

            sharedStringItem45.Append(text45);

            SharedStringItem sharedStringItem46 = new SharedStringItem();
            Text text46 = new Text();
            text46.Text = "海珠";

            sharedStringItem46.Append(text46);

            SharedStringItem sharedStringItem47 = new SharedStringItem();
            Text text47 = new Text();
            text47.Text = "2077906251";

            sharedStringItem47.Append(text47);

            SharedStringItem sharedStringItem48 = new SharedStringItem();
            Text text48 = new Text();
            text48.Text = "贡茶(海珠店)";

            sharedStringItem48.Append(text48);

            SharedStringItem sharedStringItem49 = new SharedStringItem();
            Text text49 = new Text();
            text49.Text = "海南";

            sharedStringItem49.Append(text49);

            SharedStringItem sharedStringItem50 = new SharedStringItem();
            Text text50 = new Text();
            text50.Text = "于松民";

            sharedStringItem50.Append(text50);

            SharedStringItem sharedStringItem51 = new SharedStringItem();
            Text text51 = new Text();
            text51.Text = "海口";

            sharedStringItem51.Append(text51);

            SharedStringItem sharedStringItem52 = new SharedStringItem();
            Text text52 = new Text();
            text52.Text = "10854598";

            sharedStringItem52.Append(text52);

            SharedStringItem sharedStringItem53 = new SharedStringItem();
            Text text53 = new Text();
            text53.Text = "喜三德甜品·手工芋圆（海口店）";

            sharedStringItem53.Append(text53);

            SharedStringItem sharedStringItem54 = new SharedStringItem();
            Text text54 = new Text();
            text54.Text = "厦门";

            sharedStringItem54.Append(text54);

            SharedStringItem sharedStringItem55 = new SharedStringItem();
            Text text55 = new Text();
            text55.Text = "2077997044";

            sharedStringItem55.Append(text55);

            SharedStringItem sharedStringItem56 = new SharedStringItem();
            Text text56 = new Text();
            text56.Text = "喜三德甜品●手工芋圆(厦门店)";

            sharedStringItem56.Append(text56);

            SharedStringItem sharedStringItem57 = new SharedStringItem();
            Text text57 = new Text();
            text57.Text = "深圳";

            sharedStringItem57.Append(text57);

            SharedStringItem sharedStringItem58 = new SharedStringItem();
            Text text58 = new Text();
            text58.Text = "刘文君";

            sharedStringItem58.Append(text58);

            SharedStringItem sharedStringItem59 = new SharedStringItem();
            Text text59 = new Text();
            text59.Text = "新生";

            sharedStringItem59.Append(text59);

            SharedStringItem sharedStringItem60 = new SharedStringItem();
            Text text60 = new Text();
            text60.Text = "501849656";

            sharedStringItem60.Append(text60);

            SharedStringItem sharedStringItem61 = new SharedStringItem();
            Text text61 = new Text();
            text61.Text = "贡茶(龙岗店)";

            sharedStringItem61.Append(text61);

            SharedStringItem sharedStringItem62 = new SharedStringItem();
            Text text62 = new Text();
            text62.Text = "横岗";

            sharedStringItem62.Append(text62);

            SharedStringItem sharedStringItem63 = new SharedStringItem();
            Text text63 = new Text();
            text63.Text = "501809918";

            sharedStringItem63.Append(text63);

            SharedStringItem sharedStringItem64 = new SharedStringItem();
            Text text64 = new Text();
            text64.Text = "苏姐牛奶甜品世家(华西二路店)";

            sharedStringItem64.Append(text64);

            SharedStringItem sharedStringItem65 = new SharedStringItem();
            Text text65 = new Text();
            text65.Text = "Sum of 比上周";

            sharedStringItem65.Append(text65);

            SharedStringItem sharedStringItem66 = new SharedStringItem();
            Text text66 = new Text();
            text66.Text = "(blank)";

            sharedStringItem66.Append(text66);

            SharedStringItem sharedStringItem67 = new SharedStringItem();
            Text text67 = new Text();
            text67.Text = "Grand Total";

            sharedStringItem67.Append(text67);

            SharedStringItem sharedStringItem68 = new SharedStringItem();
            Text text68 = new Text();
            text68.Text = "云南旗舰店 Total";

            sharedStringItem68.Append(text68);

            SharedStringItem sharedStringItem69 = new SharedStringItem();
            Text text69 = new Text();
            text69.Text = "厦门 Total";

            sharedStringItem69.Append(text69);

            SharedStringItem sharedStringItem70 = new SharedStringItem();
            Text text70 = new Text();
            text70.Text = "新生 Total";

            sharedStringItem70.Append(text70);

            SharedStringItem sharedStringItem71 = new SharedStringItem();
            Text text71 = new Text();
            text71.Text = "横岗 Total";

            sharedStringItem71.Append(text71);

            SharedStringItem sharedStringItem72 = new SharedStringItem();
            Text text72 = new Text();
            text72.Text = "江汉 Total";

            sharedStringItem72.Append(text72);

            SharedStringItem sharedStringItem73 = new SharedStringItem();
            Text text73 = new Text();
            text73.Text = "海口 Total";

            sharedStringItem73.Append(text73);

            SharedStringItem sharedStringItem74 = new SharedStringItem();
            Text text74 = new Text();
            text74.Text = "海珠 Total";

            sharedStringItem74.Append(text74);

            SharedStringItem sharedStringItem75 = new SharedStringItem();
            Text text75 = new Text();
            text75.Text = "赣州 Total";

            sharedStringItem75.Append(text75);

            SharedStringItem sharedStringItem76 = new SharedStringItem();
            Text text76 = new Text();
            text76.Text = "(blank) Total";

            sharedStringItem76.Append(text76);

            sharedStringTable1.Append(sharedStringItem1);
            sharedStringTable1.Append(sharedStringItem2);
            sharedStringTable1.Append(sharedStringItem3);
            sharedStringTable1.Append(sharedStringItem4);
            sharedStringTable1.Append(sharedStringItem5);
            sharedStringTable1.Append(sharedStringItem6);
            sharedStringTable1.Append(sharedStringItem7);
            sharedStringTable1.Append(sharedStringItem8);
            sharedStringTable1.Append(sharedStringItem9);
            sharedStringTable1.Append(sharedStringItem10);
            sharedStringTable1.Append(sharedStringItem11);
            sharedStringTable1.Append(sharedStringItem12);
            sharedStringTable1.Append(sharedStringItem13);
            sharedStringTable1.Append(sharedStringItem14);
            sharedStringTable1.Append(sharedStringItem15);
            sharedStringTable1.Append(sharedStringItem16);
            sharedStringTable1.Append(sharedStringItem17);
            sharedStringTable1.Append(sharedStringItem18);
            sharedStringTable1.Append(sharedStringItem19);
            sharedStringTable1.Append(sharedStringItem20);
            sharedStringTable1.Append(sharedStringItem21);
            sharedStringTable1.Append(sharedStringItem22);
            sharedStringTable1.Append(sharedStringItem23);
            sharedStringTable1.Append(sharedStringItem24);
            sharedStringTable1.Append(sharedStringItem25);
            sharedStringTable1.Append(sharedStringItem26);
            sharedStringTable1.Append(sharedStringItem27);
            sharedStringTable1.Append(sharedStringItem28);
            sharedStringTable1.Append(sharedStringItem29);
            sharedStringTable1.Append(sharedStringItem30);
            sharedStringTable1.Append(sharedStringItem31);
            sharedStringTable1.Append(sharedStringItem32);
            sharedStringTable1.Append(sharedStringItem33);
            sharedStringTable1.Append(sharedStringItem34);
            sharedStringTable1.Append(sharedStringItem35);
            sharedStringTable1.Append(sharedStringItem36);
            sharedStringTable1.Append(sharedStringItem37);
            sharedStringTable1.Append(sharedStringItem38);
            sharedStringTable1.Append(sharedStringItem39);
            sharedStringTable1.Append(sharedStringItem40);
            sharedStringTable1.Append(sharedStringItem41);
            sharedStringTable1.Append(sharedStringItem42);
            sharedStringTable1.Append(sharedStringItem43);
            sharedStringTable1.Append(sharedStringItem44);
            sharedStringTable1.Append(sharedStringItem45);
            sharedStringTable1.Append(sharedStringItem46);
            sharedStringTable1.Append(sharedStringItem47);
            sharedStringTable1.Append(sharedStringItem48);
            sharedStringTable1.Append(sharedStringItem49);
            sharedStringTable1.Append(sharedStringItem50);
            sharedStringTable1.Append(sharedStringItem51);
            sharedStringTable1.Append(sharedStringItem52);
            sharedStringTable1.Append(sharedStringItem53);
            sharedStringTable1.Append(sharedStringItem54);
            sharedStringTable1.Append(sharedStringItem55);
            sharedStringTable1.Append(sharedStringItem56);
            sharedStringTable1.Append(sharedStringItem57);
            sharedStringTable1.Append(sharedStringItem58);
            sharedStringTable1.Append(sharedStringItem59);
            sharedStringTable1.Append(sharedStringItem60);
            sharedStringTable1.Append(sharedStringItem61);
            sharedStringTable1.Append(sharedStringItem62);
            sharedStringTable1.Append(sharedStringItem63);
            sharedStringTable1.Append(sharedStringItem64);
            sharedStringTable1.Append(sharedStringItem65);
            sharedStringTable1.Append(sharedStringItem66);
            sharedStringTable1.Append(sharedStringItem67);
            sharedStringTable1.Append(sharedStringItem68);
            sharedStringTable1.Append(sharedStringItem69);
            sharedStringTable1.Append(sharedStringItem70);
            sharedStringTable1.Append(sharedStringItem71);
            sharedStringTable1.Append(sharedStringItem72);
            sharedStringTable1.Append(sharedStringItem73);
            sharedStringTable1.Append(sharedStringItem74);
            sharedStringTable1.Append(sharedStringItem75);
            sharedStringTable1.Append(sharedStringItem76);

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet(){ MCAttributes = new MarkupCompatibilityAttributes(){ Ignorable = "x14ac x16r2 xr" }  };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            Fonts fonts1 = new Fonts(){ Count = (UInt32Value)2U };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize(){ Val = 11D };
            Color color1 = new Color(){ Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName(){ Val = "宋体" };
            FontCharSet fontCharSet1 = new FontCharSet(){ Val = 134 };
            FontScheme fontScheme2 = new FontScheme(){ Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme2);

            Font font2 = new Font();
            FontSize fontSize2 = new FontSize(){ Val = 11D };
            Color color2 = new Color(){ Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName(){ Val = "宋体" };
            FontCharSet fontCharSet2 = new FontCharSet(){ Val = 134 };

            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontCharSet2);

            fonts1.Append(font1);
            fonts1.Append(font2);

            Fills fills1 = new Fills(){ Count = (UInt32Value)3U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill(){ PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill(){ PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill(){ PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor(){ Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor1 = new BackgroundColor(){ Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            fills1.Append(fill1);
            fills1.Append(fill2);
            fills1.Append(fill3);

            Borders borders1 = new Borders(){ Count = (UInt32Value)1U };

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

            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats(){ Count = (UInt32Value)1U };

            CellFormat cellFormat1 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            Alignment alignment1 = new Alignment(){ Vertical = VerticalAlignmentValues.Center };

            cellFormat1.Append(alignment1);

            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats(){ Count = (UInt32Value)5U };

            CellFormat cellFormat2 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };
            Alignment alignment2 = new Alignment(){ Vertical = VerticalAlignmentValues.Center };

            cellFormat2.Append(alignment2);
            CellFormat cellFormat3 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };

            CellFormat cellFormat4 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            Alignment alignment3 = new Alignment(){ Vertical = VerticalAlignmentValues.Center };

            cellFormat4.Append(alignment3);

            CellFormat cellFormat5 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, PivotButton = true };
            Alignment alignment4 = new Alignment(){ Vertical = VerticalAlignmentValues.Center };

            cellFormat5.Append(alignment4);

            CellFormat cellFormat6 = new CellFormat(){ NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true };
            Alignment alignment5 = new Alignment(){ Vertical = VerticalAlignmentValues.Center };

            cellFormat6.Append(alignment5);

            cellFormats1.Append(cellFormat2);
            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);

            CellStyles cellStyles1 = new CellStyles(){ Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle(){ Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles1.Append(cellStyle1);
            DifferentialFormats differentialFormats1 = new DifferentialFormats(){ Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles(){ Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension(){ Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles(){ DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension(){ Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles1 = new X15.TimelineStyles(){ DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);
            stylesheetExtensionList1.Append(stylesheetExtension2);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);
            stylesheet1.Append(cellStyles1);
            stylesheet1.Append(differentialFormats1);
            stylesheet1.Append(tableStyles1);
            stylesheet1.Append(stylesheetExtensionList1);

            workbookStylesPart1.Stylesheet = stylesheet1;
        }

        // Generates content of customFilePropertiesPart1.
        private void GenerateCustomFilePropertiesPart1Content(CustomFilePropertiesPart customFilePropertiesPart1)
        {
            Op.Properties properties2 = new Op.Properties();
            properties2.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

            Op.CustomDocumentProperty customDocumentProperty1 = new Op.CustomDocumentProperty(){ FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", PropertyId = 2, Name = "KSOProductBuildVer" };
            Vt.VTLPWSTR vTLPWSTR1 = new Vt.VTLPWSTR();
            vTLPWSTR1.Text = "2052-11.1.0.10314";

            customDocumentProperty1.Append(vTLPWSTR1);

            properties2.Append(customDocumentProperty1);

            customFilePropertiesPart1.Properties = properties2;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Administrator";
            document.PackageProperties.Title = "";
            document.PackageProperties.Subject = "";
            document.PackageProperties.Category = "";
            document.PackageProperties.Keywords = "";
            document.PackageProperties.Description = "";
            document.PackageProperties.ContentStatus = "";
            document.PackageProperties.Revision = "";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-03-07T10:10:38Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2021-03-07T10:19:14Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "hu zilch";
        }


    }
}
