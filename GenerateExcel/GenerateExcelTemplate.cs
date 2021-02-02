using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using System;

namespace ExcelOpenXML.GenerateExcel
{
    public partial class GenerateExcel
    {
        // Creates a SpreadsheetDocument.
        public void CreatePackage(string filePath)
        {
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                CreateParts(package);
            }
        }

        // Adds child parts and generates content of the specified part.
        private void CreateParts(SpreadsheetDocument document)
        {
            //ExtendedFilePropertiesPart extendedFilePropertiesPart1 = document.AddNewPart<ExtendedFilePropertiesPart>("rId3");
            //GenerateExtendedFilePropertiesPart1Content(extendedFilePropertiesPart1);

            int totalRegistros = 1000000;
            int registrosPorAba = 65535;
            double totalRegistrosD = 1000000;
            double registrosPorAbaD = 65535;
            int i = 0;
            int plan = 1;
            int numeroAbas = (int)Math.Ceiling(totalRegistrosD / registrosPorAbaD);
            //Alterar nesse método quantas abas serao criadas
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1, numeroAbas);

            for (i = 0; i < totalRegistros; i++)
            {
                if (i > 0 && i % registrosPorAba == 0)
                {
                    //Aba dados
                    WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId" + 10 + plan);
                    GenerateWorksheetContent(worksheetPart1, plan);
                    plan++;
                    //GenerateWorksheetPart1Content(worksheetPart1);
                }
            }
            //Sobra da divisao, monta uma aba com menos de 65535
            if (i % registrosPorAba != 0)
            {
                //Aba dados
                WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId" + 10 + plan);
                GenerateWorksheetContent(worksheetPart1, plan);
                plan++;
            }

            //WorksheetPart worksheetPart1 = workbookPart1.AddNewPart<WorksheetPart>("rId3");
            //GenerateWorksheetPart1Content(worksheetPart1);

            //CalculationChainPart calculationChainPart1 = workbookPart1.AddNewPart<CalculationChainPart>("rId7");
            //GenerateCalculationChainPart1Content(calculationChainPart1);

            //WorksheetPart worksheetPart2 = workbookPart1.AddNewPart<WorksheetPart>("rId2");
            //GenerateWorksheetPart2Content(worksheetPart2);

            //WorksheetPart worksheetPart3 = workbookPart1.AddNewPart<WorksheetPart>("rId1");
            //GenerateWorksheetPart3Content(worksheetPart3);

            SharedStringTablePart sharedStringTablePart1 = workbookPart1.AddNewPart<SharedStringTablePart>("rId6");
            GenerateSharedStringTablePart1Content(sharedStringTablePart1);

            WorkbookStylesPart workbookStylesPart1 = workbookPart1.AddNewPart<WorkbookStylesPart>("rId5");
            GenerateWorkbookStylesPart1Content(workbookStylesPart1);

            ThemePart themePart1 = workbookPart1.AddNewPart<ThemePart>("rId4");
            GenerateThemePart1Content(themePart1);

            SetPackageProperties(document);
        }

        // Generates content of extendedFilePropertiesPart1.
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
            vTLPSTR1.Text = "Planilhas";

            variant1.Append(vTLPSTR1);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "3";

            variant2.Append(vTInt321);

            vTVector1.Append(variant1);
            vTVector1.Append(variant2);

            headingPairs1.Append(vTVector1);

            Ap.TitlesOfParts titlesOfParts1 = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)3U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Plan1";
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Plan2";
            Vt.VTLPSTR vTLPSTR4 = new Vt.VTLPSTR();
            vTLPSTR4.Text = "Total Geral";

            vTVector2.Append(vTLPSTR2);
            vTVector2.Append(vTLPSTR3);
            vTVector2.Append(vTLPSTR4);

            titlesOfParts1.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate1 = new Ap.LinksUpToDate();
            linksUpToDate1.Text = "false";
            Ap.SharedDocument sharedDocument1 = new Ap.SharedDocument();
            sharedDocument1.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged1 = new Ap.HyperlinksChanged();
            hyperlinksChanged1.Text = "false";
            Ap.ApplicationVersion applicationVersion1 = new Ap.ApplicationVersion();
            applicationVersion1.Text = "14.0300";

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

        // Generates content of workbookPart1.
        private void GenerateWorkbookPart1Content(WorkbookPart workbookPart1, int numeroAbas)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            FileVersion fileVersion1 = new FileVersion() { ApplicationName = "xl", LastEdited = "5", LowestEdited = "5", BuildVersion = "9303" };
            WorkbookProperties workbookProperties1 = new WorkbookProperties() { DefaultThemeVersion = (UInt32Value)124226U };

            BookViews bookViews1 = new BookViews();
            WorkbookView workbookView1 = new WorkbookView() { XWindow = 480, YWindow = 105, WindowWidth = (UInt32Value)27795U, WindowHeight = (UInt32Value)12600U };

            bookViews1.Append(workbookView1);

            Sheets sheets1 = new Sheets();
            for (int i = 1; i <= numeroAbas; i++)
            {
                Sheet sheet1 = new Sheet() { Name = "Plan" + i.ToString(), SheetId = (UInt32Value)(uint)i, Id = "rId" + 10 + i };
                sheets1.Append(sheet1);
            }

            //Sheet sheet1 = new Sheet() { Name = "Plan1", SheetId = (UInt32Value)1U, Id = "rId1" };
            //Sheet sheet2 = new Sheet() { Name = "Plan2", SheetId = (UInt32Value)3U, Id = "rId2" };
            //Sheet sheet3 = new Sheet() { Name = "Total Geral", SheetId = (UInt32Value)2U, Id = "rId3" };

            //sheets1.Append(sheet1);
            //sheets1.Append(sheet2);
            //sheets1.Append(sheet3);
            CalculationProperties calculationProperties1 = new CalculationProperties() { CalculationId = (UInt32Value)145621U };

            workbook1.Append(fileVersion1);
            workbook1.Append(workbookProperties1);
            workbook1.Append(bookViews1);
            workbook1.Append(sheets1);
            workbook1.Append(calculationProperties1);

            workbookPart1.Workbook = workbook1;
        }

        // Generates content of worksheetPart1.
        private void GenerateWorksheetPart1Content(WorksheetPart worksheetPart1)
        {
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1:H12" };

            SheetViews sheetViews1 = new SheetViews();

            SheetView sheetView1 = new SheetView() { WorkbookViewId = (UInt32Value)0U };
            Selection selection1 = new Selection() { ActiveCell = "C12", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "C12" } };

            sheetView1.Append(selection1);

            sheetViews1.Append(sheetView1);
            SheetFormatProperties sheetFormatProperties1 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns1 = new Columns();
            Column column1 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 13.140625D, BestFit = true, CustomWidth = true };
            Column column2 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12.28515625D, BestFit = true, CustomWidth = true };
            Column column3 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 12.85546875D, BestFit = true, CustomWidth = true };
            Column column4 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 5.42578125D, BestFit = true, CustomWidth = true };
            Column column5 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 9D, BestFit = true, CustomWidth = true };
            Column column6 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 12.42578125D, BestFit = true, CustomWidth = true };
            Column column7 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 12D, CustomWidth = true };

            columns1.Append(column1);
            columns1.Append(column2);
            columns1.Append(column3);
            columns1.Append(column4);
            columns1.Append(column5);
            columns1.Append(column6);
            columns1.Append(column7);

            SheetData sheetData1 = new SheetData();

            Row row1 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 21D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell1 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue1 = new CellValue();
            cellValue1.Text = "0";

            cell1.Append(cellValue1);
            Cell cell2 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)16U };
            Cell cell3 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)16U };
            Cell cell4 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)16U };
            Cell cell5 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)16U };
            Cell cell6 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)16U };
            Cell cell7 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)16U };
            Cell cell8 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)16U };

            row1.Append(cell1);
            row1.Append(cell2);
            row1.Append(cell3);
            row1.Append(cell4);
            row1.Append(cell5);
            row1.Append(cell6);
            row1.Append(cell7);
            row1.Append(cell8);

            Row row2 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell9 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue2 = new CellValue();
            cellValue2.Text = "1";

            cell9.Append(cellValue2);
            Cell cell10 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)16U };
            Cell cell11 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)16U };
            Cell cell12 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)16U };
            Cell cell13 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)16U };
            Cell cell14 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)16U };
            Cell cell15 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)16U };
            Cell cell16 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)16U };

            row2.Append(cell9);
            row2.Append(cell10);
            row2.Append(cell11);
            row2.Append(cell12);
            row2.Append(cell13);
            row2.Append(cell14);
            row2.Append(cell15);
            row2.Append(cell16);

            Row row3 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 21D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell17 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)17U, DataType = CellValues.SharedString };
            CellValue cellValue3 = new CellValue();
            cellValue3.Text = "2";

            cell17.Append(cellValue3);
            Cell cell18 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)16U };
            Cell cell19 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)16U };
            Cell cell20 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)16U };
            Cell cell21 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)16U };
            Cell cell22 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)16U };
            Cell cell23 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)16U };
            Cell cell24 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)16U };

            row3.Append(cell17);
            row3.Append(cell18);
            row3.Append(cell19);
            row3.Append(cell20);
            row3.Append(cell21);
            row3.Append(cell22);
            row3.Append(cell23);
            row3.Append(cell24);

            Row row4 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell25 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue4 = new CellValue();
            cellValue4.Text = "3";

            cell25.Append(cellValue4);
            Cell cell26 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)1U };
            Cell cell27 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)1U };
            Cell cell28 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)1U };
            Cell cell29 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)2U };
            Cell cell30 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)2U };
            Cell cell31 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)3U };
            Cell cell32 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)3U };

            row4.Append(cell25);
            row4.Append(cell26);
            row4.Append(cell27);
            row4.Append(cell28);
            row4.Append(cell29);
            row4.Append(cell30);
            row4.Append(cell31);
            row4.Append(cell32);

            Row row5 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell33 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue5 = new CellValue();
            cellValue5.Text = "4";

            cell33.Append(cellValue5);
            Cell cell34 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)4U };
            Cell cell35 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)4U };
            Cell cell36 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)1U };
            Cell cell37 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)2U };
            Cell cell38 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)2U };
            Cell cell39 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)3U };
            Cell cell40 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)3U };

            row5.Append(cell33);
            row5.Append(cell34);
            row5.Append(cell35);
            row5.Append(cell36);
            row5.Append(cell37);
            row5.Append(cell38);
            row5.Append(cell39);
            row5.Append(cell40);

            Row row6 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell41 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)5U };
            CellValue cellValue6 = new CellValue();
            cellValue6.Text = "44166.661615277779";

            cell41.Append(cellValue6);
            Cell cell42 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)4U };
            Cell cell43 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)4U };
            Cell cell44 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)1U };
            Cell cell45 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)2U };
            Cell cell46 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)2U };
            Cell cell47 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)3U };
            Cell cell48 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)3U };

            row6.Append(cell41);
            row6.Append(cell42);
            row6.Append(cell43);
            row6.Append(cell44);
            row6.Append(cell45);
            row6.Append(cell46);
            row6.Append(cell47);
            row6.Append(cell48);

            Row row7 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell49 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue7 = new CellValue();
            cellValue7.Text = "3";

            cell49.Append(cellValue7);
            Cell cell50 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)7U };
            Cell cell51 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)8U };
            Cell cell52 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)8U };
            Cell cell53 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)9U };
            Cell cell54 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)9U };
            Cell cell55 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)8U };
            Cell cell56 = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)8U };

            row7.Append(cell49);
            row7.Append(cell50);
            row7.Append(cell51);
            row7.Append(cell52);
            row7.Append(cell53);
            row7.Append(cell54);
            row7.Append(cell55);
            row7.Append(cell56);

            Row row8 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell57 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue8 = new CellValue();
            cellValue8.Text = "5";

            cell57.Append(cellValue8);

            Cell cell58 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue9 = new CellValue();
            cellValue9.Text = "6";

            cell58.Append(cellValue9);

            Cell cell59 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue10 = new CellValue();
            cellValue10.Text = "7";

            cell59.Append(cellValue10);

            Cell cell60 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue11 = new CellValue();
            cellValue11.Text = "8";

            cell60.Append(cellValue11);

            Cell cell61 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)20U, DataType = CellValues.SharedString };
            CellValue cellValue12 = new CellValue();
            cellValue12.Text = "9";

            cell61.Append(cellValue12);

            Cell cell62 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)20U, DataType = CellValues.SharedString };
            CellValue cellValue13 = new CellValue();
            cellValue13.Text = "10";

            cell62.Append(cellValue13);

            Cell cell63 = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue14 = new CellValue();
            cellValue14.Text = "11";

            cell63.Append(cellValue14);

            Cell cell64 = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue15 = new CellValue();
            cellValue15.Text = "12";

            cell64.Append(cellValue15);

            row8.Append(cell57);
            row8.Append(cell58);
            row8.Append(cell59);
            row8.Append(cell60);
            row8.Append(cell61);
            row8.Append(cell62);
            row8.Append(cell63);
            row8.Append(cell64);

            Row row9 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
            Cell cell65 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)19U };
            Cell cell66 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)19U };
            Cell cell67 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)19U };
            Cell cell68 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)19U };
            Cell cell69 = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)21U };
            Cell cell70 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)21U };
            Cell cell71 = new Cell() { CellReference = "G9", StyleIndex = (UInt32Value)14U };
            Cell cell72 = new Cell() { CellReference = "H9", StyleIndex = (UInt32Value)14U };

            row9.Append(cell65);
            row9.Append(cell66);
            row9.Append(cell67);
            row9.Append(cell68);
            row9.Append(cell69);
            row9.Append(cell70);
            row9.Append(cell71);
            row9.Append(cell72);

            Row row10 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell73 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue16 = new CellValue();
            cellValue16.Text = "13";

            cell73.Append(cellValue16);

            Cell cell74 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula1 = new CellFormula();
            cellFormula1.Text = "Plan1!B11";
            CellValue cellValue17 = new CellValue();
            cellValue17.Text = "7812.89";

            cell74.Append(cellFormula1);
            cell74.Append(cellValue17);

            Cell cell75 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula2 = new CellFormula();
            cellFormula2.Text = "Plan1!C11";
            CellValue cellValue18 = new CellValue();
            cellValue18.Text = "7823.63";

            cell75.Append(cellFormula2);
            cell75.Append(cellValue18);

            Cell cell76 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula3 = new CellFormula();
            cellFormula3.Text = "Plan1!D11";
            CellValue cellValue19 = new CellValue();
            cellValue19.Text = "-";

            cell76.Append(cellFormula3);
            cell76.Append(cellValue19);

            Cell cell77 = new Cell() { CellReference = "E10", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula4 = new CellFormula();
            cellFormula4.Text = "Plan1!E11";
            CellValue cellValue20 = new CellValue();
            cellValue20.Text = "-";

            cell77.Append(cellFormula4);
            cell77.Append(cellValue20);

            Cell cell78 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula5 = new CellFormula();
            cellFormula5.Text = "Plan1!F11";
            CellValue cellValue21 = new CellValue();
            cellValue21.Text = "-";

            cell78.Append(cellFormula5);
            cell78.Append(cellValue21);

            Cell cell79 = new Cell() { CellReference = "G10", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula6 = new CellFormula();
            cellFormula6.Text = "Plan1!G11";
            CellValue cellValue22 = new CellValue();
            cellValue22.Text = "-";

            cell79.Append(cellFormula6);
            cell79.Append(cellValue22);

            Cell cell80 = new Cell() { CellReference = "H10", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula7 = new CellFormula();
            cellFormula7.Text = "Plan1!H11";
            CellValue cellValue23 = new CellValue();
            cellValue23.Text = "-";

            cell80.Append(cellFormula7);
            cell80.Append(cellValue23);

            row10.Append(cell73);
            row10.Append(cell74);
            row10.Append(cell75);
            row10.Append(cell76);
            row10.Append(cell77);
            row10.Append(cell78);
            row10.Append(cell79);
            row10.Append(cell80);

            Row row11 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell81 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue24 = new CellValue();
            cellValue24.Text = "20";

            cell81.Append(cellValue24);

            Cell cell82 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula8 = new CellFormula();
            cellFormula8.Text = "Plan2!B11";
            CellValue cellValue25 = new CellValue();
            cellValue25.Text = "8974.5400000000009";

            cell82.Append(cellFormula8);
            cell82.Append(cellValue25);

            Cell cell83 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula9 = new CellFormula();
            cellFormula9.Text = "Plan2!C11";
            CellValue cellValue26 = new CellValue();
            cellValue26.Text = "2354.44";

            cell83.Append(cellFormula9);
            cell83.Append(cellValue26);

            Cell cell84 = new Cell() { CellReference = "D11", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula10 = new CellFormula();
            cellFormula10.Text = "Plan2!D11";
            CellValue cellValue27 = new CellValue();
            cellValue27.Text = "-";

            cell84.Append(cellFormula10);
            cell84.Append(cellValue27);

            Cell cell85 = new Cell() { CellReference = "E11", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula11 = new CellFormula();
            cellFormula11.Text = "Plan2!E11";
            CellValue cellValue28 = new CellValue();
            cellValue28.Text = "-";

            cell85.Append(cellFormula11);
            cell85.Append(cellValue28);

            Cell cell86 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula12 = new CellFormula();
            cellFormula12.Text = "Plan2!F11";
            CellValue cellValue29 = new CellValue();
            cellValue29.Text = "-";

            cell86.Append(cellFormula12);
            cell86.Append(cellValue29);

            Cell cell87 = new Cell() { CellReference = "G11", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula13 = new CellFormula();
            cellFormula13.Text = "Plan2!G11";
            CellValue cellValue30 = new CellValue();
            cellValue30.Text = "-";

            cell87.Append(cellFormula13);
            cell87.Append(cellValue30);

            Cell cell88 = new Cell() { CellReference = "H11", StyleIndex = (UInt32Value)10U, DataType = CellValues.String };
            CellFormula cellFormula14 = new CellFormula();
            cellFormula14.Text = "Plan2!H11";
            CellValue cellValue31 = new CellValue();
            cellValue31.Text = "-";

            cell88.Append(cellFormula14);
            cell88.Append(cellValue31);

            row11.Append(cell81);
            row11.Append(cell82);
            row11.Append(cell83);
            row11.Append(cell84);
            row11.Append(cell85);
            row11.Append(cell86);
            row11.Append(cell87);
            row11.Append(cell88);

            Row row12 = new Row() { RowIndex = (UInt32Value)12U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell89 = new Cell() { CellReference = "A12", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue32 = new CellValue();
            cellValue32.Text = "16";

            cell89.Append(cellValue32);

            Cell cell90 = new Cell() { CellReference = "B12", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula15 = new CellFormula();
            cellFormula15.Text = "SUM(B10:B11)";
            CellValue cellValue33 = new CellValue();
            cellValue33.Text = "16787.43";

            cell90.Append(cellFormula15);
            cell90.Append(cellValue33);

            Cell cell91 = new Cell() { CellReference = "C12", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula16 = new CellFormula();
            cellFormula16.Text = "SUM(C10:C11)";
            CellValue cellValue34 = new CellValue();
            cellValue34.Text = "10178.07";

            cell91.Append(cellFormula16);
            cell91.Append(cellValue34);

            Cell cell92 = new Cell() { CellReference = "D12", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue35 = new CellValue();
            cellValue35.Text = "17";

            cell92.Append(cellValue35);

            Cell cell93 = new Cell() { CellReference = "E12", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue36 = new CellValue();
            cellValue36.Text = "17";

            cell93.Append(cellValue36);

            Cell cell94 = new Cell() { CellReference = "F12", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue37 = new CellValue();
            cellValue37.Text = "17";

            cell94.Append(cellValue37);

            Cell cell95 = new Cell() { CellReference = "G12", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue38 = new CellValue();
            cellValue38.Text = "17";

            cell95.Append(cellValue38);

            Cell cell96 = new Cell() { CellReference = "H12", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue39 = new CellValue();
            cellValue39.Text = "17";

            cell96.Append(cellValue39);

            row12.Append(cell89);
            row12.Append(cell90);
            row12.Append(cell91);
            row12.Append(cell92);
            row12.Append(cell93);
            row12.Append(cell94);
            row12.Append(cell95);
            row12.Append(cell96);

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

            MergeCells mergeCells1 = new MergeCells() { Count = (UInt32Value)11U };
            MergeCell mergeCell1 = new MergeCell() { Reference = "H8:H9" };
            MergeCell mergeCell2 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell3 = new MergeCell() { Reference = "A2:H2" };
            MergeCell mergeCell4 = new MergeCell() { Reference = "A3:H3" };
            MergeCell mergeCell5 = new MergeCell() { Reference = "A8:A9" };
            MergeCell mergeCell6 = new MergeCell() { Reference = "B8:B9" };
            MergeCell mergeCell7 = new MergeCell() { Reference = "C8:C9" };
            MergeCell mergeCell8 = new MergeCell() { Reference = "D8:D9" };
            MergeCell mergeCell9 = new MergeCell() { Reference = "E8:E9" };
            MergeCell mergeCell10 = new MergeCell() { Reference = "F8:F9" };
            MergeCell mergeCell11 = new MergeCell() { Reference = "G8:G9" };

            mergeCells1.Append(mergeCell1);
            mergeCells1.Append(mergeCell2);
            mergeCells1.Append(mergeCell3);
            mergeCells1.Append(mergeCell4);
            mergeCells1.Append(mergeCell5);
            mergeCells1.Append(mergeCell6);
            mergeCells1.Append(mergeCell7);
            mergeCells1.Append(mergeCell8);
            mergeCells1.Append(mergeCell9);
            mergeCells1.Append(mergeCell10);
            mergeCells1.Append(mergeCell11);
            PageMargins pageMargins1 = new PageMargins() { Left = 0.511811024D, Right = 0.511811024D, Top = 0.78740157499999996D, Bottom = 0.78740157499999996D, Header = 0.31496062000000002D, Footer = 0.31496062000000002D };

            worksheet1.Append(sheetDimension1);
            worksheet1.Append(sheetViews1);
            worksheet1.Append(sheetFormatProperties1);
            worksheet1.Append(columns1);
            worksheet1.Append(sheetData1);
            worksheet1.Append(mergeCells1);
            worksheet1.Append(pageMargins1);

            worksheetPart1.Worksheet = worksheet1;
        }

        // Generates content of calculationChainPart1.
        private void GenerateCalculationChainPart1Content(CalculationChainPart calculationChainPart1)
        {
            CalculationChain calculationChain1 = new CalculationChain();
            CalculationCell calculationCell1 = new CalculationCell() { CellReference = "C12", SheetId = 2, NewLevel = true };
            CalculationCell calculationCell2 = new CalculationCell() { CellReference = "B12", SheetId = 2 };
            CalculationCell calculationCell3 = new CalculationCell() { CellReference = "C11", SheetId = 2 };
            CalculationCell calculationCell4 = new CalculationCell() { CellReference = "D11", SheetId = 2 };
            CalculationCell calculationCell5 = new CalculationCell() { CellReference = "E11", SheetId = 2 };
            CalculationCell calculationCell6 = new CalculationCell() { CellReference = "F11", SheetId = 2 };
            CalculationCell calculationCell7 = new CalculationCell() { CellReference = "G11", SheetId = 2 };
            CalculationCell calculationCell8 = new CalculationCell() { CellReference = "H11", SheetId = 2 };
            CalculationCell calculationCell9 = new CalculationCell() { CellReference = "B11", SheetId = 2 };
            CalculationCell calculationCell10 = new CalculationCell() { CellReference = "C11", SheetId = 3 };
            CalculationCell calculationCell11 = new CalculationCell() { CellReference = "B11", SheetId = 3 };
            CalculationCell calculationCell12 = new CalculationCell() { CellReference = "D10", SheetId = 2, NewLevel = true };
            CalculationCell calculationCell13 = new CalculationCell() { CellReference = "E10", SheetId = 2 };
            CalculationCell calculationCell14 = new CalculationCell() { CellReference = "F10", SheetId = 2 };
            CalculationCell calculationCell15 = new CalculationCell() { CellReference = "G10", SheetId = 2 };
            CalculationCell calculationCell16 = new CalculationCell() { CellReference = "H10", SheetId = 2 };
            CalculationCell calculationCell17 = new CalculationCell() { CellReference = "C10", SheetId = 2 };
            CalculationCell calculationCell18 = new CalculationCell() { CellReference = "B10", SheetId = 2 };
            CalculationCell calculationCell19 = new CalculationCell() { CellReference = "C11", SheetId = 1 };
            CalculationCell calculationCell20 = new CalculationCell() { CellReference = "B11", SheetId = 1 };

            calculationChain1.Append(calculationCell1);
            calculationChain1.Append(calculationCell2);
            calculationChain1.Append(calculationCell3);
            calculationChain1.Append(calculationCell4);
            calculationChain1.Append(calculationCell5);
            calculationChain1.Append(calculationCell6);
            calculationChain1.Append(calculationCell7);
            calculationChain1.Append(calculationCell8);
            calculationChain1.Append(calculationCell9);
            calculationChain1.Append(calculationCell10);
            calculationChain1.Append(calculationCell11);
            calculationChain1.Append(calculationCell12);
            calculationChain1.Append(calculationCell13);
            calculationChain1.Append(calculationCell14);
            calculationChain1.Append(calculationCell15);
            calculationChain1.Append(calculationCell16);
            calculationChain1.Append(calculationCell17);
            calculationChain1.Append(calculationCell18);
            calculationChain1.Append(calculationCell19);
            calculationChain1.Append(calculationCell20);

            calculationChainPart1.CalculationChain = calculationChain1;
        }

        // Generates content of worksheetPart2.
        private void GenerateWorksheetPart2Content(WorksheetPart worksheetPart2)
        {
            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension2 = new SheetDimension() { Reference = "A1:H11" };

            SheetViews sheetViews2 = new SheetViews();

            SheetView sheetView2 = new SheetView() { WorkbookViewId = (UInt32Value)0U };
            Selection selection2 = new Selection() { ActiveCell = "H11", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "H11" } };

            sheetView2.Append(selection2);

            sheetViews2.Append(sheetView2);
            SheetFormatProperties sheetFormatProperties2 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns2 = new Columns();
            Column column8 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 13.140625D, BestFit = true, CustomWidth = true };
            Column column9 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12.28515625D, BestFit = true, CustomWidth = true };
            Column column10 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 12.85546875D, BestFit = true, CustomWidth = true };
            Column column11 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 5.42578125D, BestFit = true, CustomWidth = true };
            Column column12 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 9D, BestFit = true, CustomWidth = true };
            Column column13 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 12.42578125D, BestFit = true, CustomWidth = true };
            Column column14 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 18.28515625D, CustomWidth = true };

            columns2.Append(column8);
            columns2.Append(column9);
            columns2.Append(column10);
            columns2.Append(column11);
            columns2.Append(column12);
            columns2.Append(column13);
            columns2.Append(column14);

            SheetData sheetData2 = new SheetData();

            Row row13 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 25.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell97 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue40 = new CellValue();
            cellValue40.Text = "0";

            cell97.Append(cellValue40);
            Cell cell98 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)16U };
            Cell cell99 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)16U };
            Cell cell100 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)16U };
            Cell cell101 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)16U };
            Cell cell102 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)16U };
            Cell cell103 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)16U };
            Cell cell104 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)16U };

            row13.Append(cell97);
            row13.Append(cell98);
            row13.Append(cell99);
            row13.Append(cell100);
            row13.Append(cell101);
            row13.Append(cell102);
            row13.Append(cell103);
            row13.Append(cell104);

            Row row14 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 21.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell105 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue41 = new CellValue();
            cellValue41.Text = "1";

            cell105.Append(cellValue41);
            Cell cell106 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)16U };
            Cell cell107 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)16U };
            Cell cell108 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)16U };
            Cell cell109 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)16U };
            Cell cell110 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)16U };
            Cell cell111 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)16U };
            Cell cell112 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)16U };

            row14.Append(cell105);
            row14.Append(cell106);
            row14.Append(cell107);
            row14.Append(cell108);
            row14.Append(cell109);
            row14.Append(cell110);
            row14.Append(cell111);
            row14.Append(cell112);

            Row row15 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 21.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell113 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)17U, DataType = CellValues.SharedString };
            CellValue cellValue42 = new CellValue();
            cellValue42.Text = "2";

            cell113.Append(cellValue42);
            Cell cell114 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)16U };
            Cell cell115 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)16U };
            Cell cell116 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)16U };
            Cell cell117 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)16U };
            Cell cell118 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)16U };
            Cell cell119 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)16U };
            Cell cell120 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)16U };

            row15.Append(cell113);
            row15.Append(cell114);
            row15.Append(cell115);
            row15.Append(cell116);
            row15.Append(cell117);
            row15.Append(cell118);
            row15.Append(cell119);
            row15.Append(cell120);

            Row row16 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell121 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue43 = new CellValue();
            cellValue43.Text = "3";

            cell121.Append(cellValue43);
            Cell cell122 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)1U };
            Cell cell123 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)1U };
            Cell cell124 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)1U };
            Cell cell125 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)2U };
            Cell cell126 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)2U };
            Cell cell127 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)3U };
            Cell cell128 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)3U };

            row16.Append(cell121);
            row16.Append(cell122);
            row16.Append(cell123);
            row16.Append(cell124);
            row16.Append(cell125);
            row16.Append(cell126);
            row16.Append(cell127);
            row16.Append(cell128);

            Row row17 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell129 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue44 = new CellValue();
            cellValue44.Text = "4";

            cell129.Append(cellValue44);
            Cell cell130 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)4U };
            Cell cell131 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)4U };
            Cell cell132 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)1U };
            Cell cell133 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)2U };
            Cell cell134 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)2U };
            Cell cell135 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)3U };
            Cell cell136 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)3U };

            row17.Append(cell129);
            row17.Append(cell130);
            row17.Append(cell131);
            row17.Append(cell132);
            row17.Append(cell133);
            row17.Append(cell134);
            row17.Append(cell135);
            row17.Append(cell136);

            Row row18 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell137 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)5U };
            CellValue cellValue45 = new CellValue();
            cellValue45.Text = "44166.661615277779";

            cell137.Append(cellValue45);
            Cell cell138 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)4U };
            Cell cell139 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)4U };
            Cell cell140 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)1U };
            Cell cell141 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)2U };
            Cell cell142 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)2U };
            Cell cell143 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)3U };
            Cell cell144 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)3U };

            row18.Append(cell137);
            row18.Append(cell138);
            row18.Append(cell139);
            row18.Append(cell140);
            row18.Append(cell141);
            row18.Append(cell142);
            row18.Append(cell143);
            row18.Append(cell144);

            Row row19 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell145 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue46 = new CellValue();
            cellValue46.Text = "3";

            cell145.Append(cellValue46);
            Cell cell146 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)7U };
            Cell cell147 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)8U };
            Cell cell148 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)8U };
            Cell cell149 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)9U };
            Cell cell150 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)9U };
            Cell cell151 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)8U };
            Cell cell152 = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)8U };

            row19.Append(cell145);
            row19.Append(cell146);
            row19.Append(cell147);
            row19.Append(cell148);
            row19.Append(cell149);
            row19.Append(cell150);
            row19.Append(cell151);
            row19.Append(cell152);

            Row row20 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell153 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue47 = new CellValue();
            cellValue47.Text = "5";

            cell153.Append(cellValue47);

            Cell cell154 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue48 = new CellValue();
            cellValue48.Text = "6";

            cell154.Append(cellValue48);

            Cell cell155 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue49 = new CellValue();
            cellValue49.Text = "7";

            cell155.Append(cellValue49);

            Cell cell156 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue50 = new CellValue();
            cellValue50.Text = "8";

            cell156.Append(cellValue50);

            Cell cell157 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)20U, DataType = CellValues.SharedString };
            CellValue cellValue51 = new CellValue();
            cellValue51.Text = "9";

            cell157.Append(cellValue51);

            Cell cell158 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)20U, DataType = CellValues.SharedString };
            CellValue cellValue52 = new CellValue();
            cellValue52.Text = "10";

            cell158.Append(cellValue52);

            Cell cell159 = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue53 = new CellValue();
            cellValue53.Text = "11";

            cell159.Append(cellValue53);

            Cell cell160 = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue54 = new CellValue();
            cellValue54.Text = "12";

            cell160.Append(cellValue54);

            row20.Append(cell153);
            row20.Append(cell154);
            row20.Append(cell155);
            row20.Append(cell156);
            row20.Append(cell157);
            row20.Append(cell158);
            row20.Append(cell159);
            row20.Append(cell160);

            Row row21 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
            Cell cell161 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)19U };
            Cell cell162 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)19U };
            Cell cell163 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)19U };
            Cell cell164 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)19U };
            Cell cell165 = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)21U };
            Cell cell166 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)21U };
            Cell cell167 = new Cell() { CellReference = "G9", StyleIndex = (UInt32Value)14U };
            Cell cell168 = new Cell() { CellReference = "H9", StyleIndex = (UInt32Value)14U };

            row21.Append(cell161);
            row21.Append(cell162);
            row21.Append(cell163);
            row21.Append(cell164);
            row21.Append(cell165);
            row21.Append(cell166);
            row21.Append(cell167);
            row21.Append(cell168);

            Row row22 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell169 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue55 = new CellValue();
            cellValue55.Text = "13";

            cell169.Append(cellValue55);

            Cell cell170 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue56 = new CellValue();
            cellValue56.Text = "8974.5400000000009";

            cell170.Append(cellValue56);

            Cell cell171 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue57 = new CellValue();
            cellValue57.Text = "2354.44";

            cell171.Append(cellValue57);

            Cell cell172 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue58 = new CellValue();
            cellValue58.Text = "3654";

            cell172.Append(cellValue58);

            Cell cell173 = new Cell() { CellReference = "E10", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue59 = new CellValue();
            cellValue59.Text = "18";

            cell173.Append(cellValue59);

            Cell cell174 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue60 = new CellValue();
            cellValue60.Text = "19";

            cell174.Append(cellValue60);

            Cell cell175 = new Cell() { CellReference = "G10", StyleIndex = (UInt32Value)12U };
            CellValue cellValue61 = new CellValue();
            cellValue61.Text = "1.01";

            cell175.Append(cellValue61);

            Cell cell176 = new Cell() { CellReference = "H10", StyleIndex = (UInt32Value)12U };
            CellValue cellValue62 = new CellValue();
            cellValue62.Text = "0.21";

            cell176.Append(cellValue62);

            row22.Append(cell169);
            row22.Append(cell170);
            row22.Append(cell171);
            row22.Append(cell172);
            row22.Append(cell173);
            row22.Append(cell174);
            row22.Append(cell175);
            row22.Append(cell176);

            Row row23 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell177 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue63 = new CellValue();
            cellValue63.Text = "16";

            cell177.Append(cellValue63);

            Cell cell178 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula17 = new CellFormula();
            cellFormula17.Text = "SUM(B10:B10)";
            CellValue cellValue64 = new CellValue();
            cellValue64.Text = "8974.5400000000009";

            cell178.Append(cellFormula17);
            cell178.Append(cellValue64);

            Cell cell179 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula18 = new CellFormula();
            cellFormula18.Text = "SUM(C10:C10)";
            CellValue cellValue65 = new CellValue();
            cellValue65.Text = "2354.44";

            cell179.Append(cellFormula18);
            cell179.Append(cellValue65);

            Cell cell180 = new Cell() { CellReference = "D11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue66 = new CellValue();
            cellValue66.Text = "17";

            cell180.Append(cellValue66);

            Cell cell181 = new Cell() { CellReference = "E11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue67 = new CellValue();
            cellValue67.Text = "17";

            cell181.Append(cellValue67);

            Cell cell182 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue68 = new CellValue();
            cellValue68.Text = "17";

            cell182.Append(cellValue68);

            Cell cell183 = new Cell() { CellReference = "G11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue69 = new CellValue();
            cellValue69.Text = "17";

            cell183.Append(cellValue69);

            Cell cell184 = new Cell() { CellReference = "H11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue70 = new CellValue();
            cellValue70.Text = "17";

            cell184.Append(cellValue70);

            row23.Append(cell177);
            row23.Append(cell178);
            row23.Append(cell179);
            row23.Append(cell180);
            row23.Append(cell181);
            row23.Append(cell182);
            row23.Append(cell183);
            row23.Append(cell184);

            sheetData2.Append(row13);
            sheetData2.Append(row14);
            sheetData2.Append(row15);
            sheetData2.Append(row16);
            sheetData2.Append(row17);
            sheetData2.Append(row18);
            sheetData2.Append(row19);
            sheetData2.Append(row20);
            sheetData2.Append(row21);
            sheetData2.Append(row22);
            sheetData2.Append(row23);

            MergeCells mergeCells2 = new MergeCells() { Count = (UInt32Value)11U };
            MergeCell mergeCell12 = new MergeCell() { Reference = "H8:H9" };
            MergeCell mergeCell13 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell14 = new MergeCell() { Reference = "A2:H2" };
            MergeCell mergeCell15 = new MergeCell() { Reference = "A3:H3" };
            MergeCell mergeCell16 = new MergeCell() { Reference = "A8:A9" };
            MergeCell mergeCell17 = new MergeCell() { Reference = "B8:B9" };
            MergeCell mergeCell18 = new MergeCell() { Reference = "C8:C9" };
            MergeCell mergeCell19 = new MergeCell() { Reference = "D8:D9" };
            MergeCell mergeCell20 = new MergeCell() { Reference = "E8:E9" };
            MergeCell mergeCell21 = new MergeCell() { Reference = "F8:F9" };
            MergeCell mergeCell22 = new MergeCell() { Reference = "G8:G9" };

            mergeCells2.Append(mergeCell12);
            mergeCells2.Append(mergeCell13);
            mergeCells2.Append(mergeCell14);
            mergeCells2.Append(mergeCell15);
            mergeCells2.Append(mergeCell16);
            mergeCells2.Append(mergeCell17);
            mergeCells2.Append(mergeCell18);
            mergeCells2.Append(mergeCell19);
            mergeCells2.Append(mergeCell20);
            mergeCells2.Append(mergeCell21);
            mergeCells2.Append(mergeCell22);
            PageMargins pageMargins2 = new PageMargins() { Left = 0.511811024D, Right = 0.511811024D, Top = 0.78740157499999996D, Bottom = 0.78740157499999996D, Header = 0.31496062000000002D, Footer = 0.31496062000000002D };

            worksheet2.Append(sheetDimension2);
            worksheet2.Append(sheetViews2);
            worksheet2.Append(sheetFormatProperties2);
            worksheet2.Append(columns2);
            worksheet2.Append(sheetData2);
            worksheet2.Append(mergeCells2);
            worksheet2.Append(pageMargins2);

            worksheetPart2.Worksheet = worksheet2;
        }

        // Generates content of worksheetPart3.
        private void GenerateWorksheetPart3Content(WorksheetPart worksheetPart3)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:H11" };

            SheetViews sheetViews3 = new SheetViews();

            SheetView sheetView3 = new SheetView() { TabSelected = true, WorkbookViewId = (UInt32Value)0U };
            Selection selection3 = new Selection() { ActiveCell = "G22", SequenceOfReferences = new ListValue<StringValue>() { InnerText = "G22" } };

            sheetView3.Append(selection3);

            sheetViews3.Append(sheetView3);
            SheetFormatProperties sheetFormatProperties3 = new SheetFormatProperties() { DefaultRowHeight = 15D, DyDescent = 0.25D };

            Columns columns3 = new Columns();
            Column column15 = new Column() { Min = (UInt32Value)1U, Max = (UInt32Value)1U, Width = 13.140625D, BestFit = true, CustomWidth = true };
            Column column16 = new Column() { Min = (UInt32Value)2U, Max = (UInt32Value)2U, Width = 12.28515625D, BestFit = true, CustomWidth = true };
            Column column17 = new Column() { Min = (UInt32Value)3U, Max = (UInt32Value)3U, Width = 12.85546875D, BestFit = true, CustomWidth = true };
            Column column18 = new Column() { Min = (UInt32Value)4U, Max = (UInt32Value)4U, Width = 5.42578125D, BestFit = true, CustomWidth = true };
            Column column19 = new Column() { Min = (UInt32Value)5U, Max = (UInt32Value)6U, Width = 9D, BestFit = true, CustomWidth = true };
            Column column20 = new Column() { Min = (UInt32Value)7U, Max = (UInt32Value)7U, Width = 12.42578125D, BestFit = true, CustomWidth = true };
            Column column21 = new Column() { Min = (UInt32Value)8U, Max = (UInt32Value)8U, Width = 18.28515625D, CustomWidth = true };

            columns3.Append(column15);
            columns3.Append(column16);
            columns3.Append(column17);
            columns3.Append(column18);
            columns3.Append(column19);
            columns3.Append(column20);
            columns3.Append(column21);

            SheetData sheetData3 = new SheetData();

            Row row24 = new Row() { RowIndex = (UInt32Value)1U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 25.5D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell185 = new Cell() { CellReference = "A1", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue71 = new CellValue();
            cellValue71.Text = "0";

            cell185.Append(cellValue71);
            Cell cell186 = new Cell() { CellReference = "B1", StyleIndex = (UInt32Value)16U };
            Cell cell187 = new Cell() { CellReference = "C1", StyleIndex = (UInt32Value)16U };
            Cell cell188 = new Cell() { CellReference = "D1", StyleIndex = (UInt32Value)16U };
            Cell cell189 = new Cell() { CellReference = "E1", StyleIndex = (UInt32Value)16U };
            Cell cell190 = new Cell() { CellReference = "F1", StyleIndex = (UInt32Value)16U };
            Cell cell191 = new Cell() { CellReference = "G1", StyleIndex = (UInt32Value)16U };
            Cell cell192 = new Cell() { CellReference = "H1", StyleIndex = (UInt32Value)16U };

            row24.Append(cell185);
            row24.Append(cell186);
            row24.Append(cell187);
            row24.Append(cell188);
            row24.Append(cell189);
            row24.Append(cell190);
            row24.Append(cell191);
            row24.Append(cell192);

            Row row25 = new Row() { RowIndex = (UInt32Value)2U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 21.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell193 = new Cell() { CellReference = "A2", StyleIndex = (UInt32Value)15U, DataType = CellValues.SharedString };
            CellValue cellValue72 = new CellValue();
            cellValue72.Text = "1";

            cell193.Append(cellValue72);
            Cell cell194 = new Cell() { CellReference = "B2", StyleIndex = (UInt32Value)16U };
            Cell cell195 = new Cell() { CellReference = "C2", StyleIndex = (UInt32Value)16U };
            Cell cell196 = new Cell() { CellReference = "D2", StyleIndex = (UInt32Value)16U };
            Cell cell197 = new Cell() { CellReference = "E2", StyleIndex = (UInt32Value)16U };
            Cell cell198 = new Cell() { CellReference = "F2", StyleIndex = (UInt32Value)16U };
            Cell cell199 = new Cell() { CellReference = "G2", StyleIndex = (UInt32Value)16U };
            Cell cell200 = new Cell() { CellReference = "H2", StyleIndex = (UInt32Value)16U };

            row25.Append(cell193);
            row25.Append(cell194);
            row25.Append(cell195);
            row25.Append(cell196);
            row25.Append(cell197);
            row25.Append(cell198);
            row25.Append(cell199);
            row25.Append(cell200);

            Row row26 = new Row() { RowIndex = (UInt32Value)3U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 21.75D, CustomHeight = true, DyDescent = 0.25D };

            Cell cell201 = new Cell() { CellReference = "A3", StyleIndex = (UInt32Value)17U, DataType = CellValues.SharedString };
            CellValue cellValue73 = new CellValue();
            cellValue73.Text = "2";

            cell201.Append(cellValue73);
            Cell cell202 = new Cell() { CellReference = "B3", StyleIndex = (UInt32Value)16U };
            Cell cell203 = new Cell() { CellReference = "C3", StyleIndex = (UInt32Value)16U };
            Cell cell204 = new Cell() { CellReference = "D3", StyleIndex = (UInt32Value)16U };
            Cell cell205 = new Cell() { CellReference = "E3", StyleIndex = (UInt32Value)16U };
            Cell cell206 = new Cell() { CellReference = "F3", StyleIndex = (UInt32Value)16U };
            Cell cell207 = new Cell() { CellReference = "G3", StyleIndex = (UInt32Value)16U };
            Cell cell208 = new Cell() { CellReference = "H3", StyleIndex = (UInt32Value)16U };

            row26.Append(cell201);
            row26.Append(cell202);
            row26.Append(cell203);
            row26.Append(cell204);
            row26.Append(cell205);
            row26.Append(cell206);
            row26.Append(cell207);
            row26.Append(cell208);

            Row row27 = new Row() { RowIndex = (UInt32Value)4U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell209 = new Cell() { CellReference = "A4", StyleIndex = (UInt32Value)1U, DataType = CellValues.SharedString };
            CellValue cellValue74 = new CellValue();
            cellValue74.Text = "3";

            cell209.Append(cellValue74);
            Cell cell210 = new Cell() { CellReference = "B4", StyleIndex = (UInt32Value)1U };
            Cell cell211 = new Cell() { CellReference = "C4", StyleIndex = (UInt32Value)1U };
            Cell cell212 = new Cell() { CellReference = "D4", StyleIndex = (UInt32Value)1U };
            Cell cell213 = new Cell() { CellReference = "E4", StyleIndex = (UInt32Value)2U };
            Cell cell214 = new Cell() { CellReference = "F4", StyleIndex = (UInt32Value)2U };
            Cell cell215 = new Cell() { CellReference = "G4", StyleIndex = (UInt32Value)3U };
            Cell cell216 = new Cell() { CellReference = "H4", StyleIndex = (UInt32Value)3U };

            row27.Append(cell209);
            row27.Append(cell210);
            row27.Append(cell211);
            row27.Append(cell212);
            row27.Append(cell213);
            row27.Append(cell214);
            row27.Append(cell215);
            row27.Append(cell216);

            Row row28 = new Row() { RowIndex = (UInt32Value)5U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell217 = new Cell() { CellReference = "A5", StyleIndex = (UInt32Value)4U, DataType = CellValues.SharedString };
            CellValue cellValue75 = new CellValue();
            cellValue75.Text = "4";

            cell217.Append(cellValue75);
            Cell cell218 = new Cell() { CellReference = "B5", StyleIndex = (UInt32Value)4U };
            Cell cell219 = new Cell() { CellReference = "C5", StyleIndex = (UInt32Value)4U };
            Cell cell220 = new Cell() { CellReference = "D5", StyleIndex = (UInt32Value)1U };
            Cell cell221 = new Cell() { CellReference = "E5", StyleIndex = (UInt32Value)2U };
            Cell cell222 = new Cell() { CellReference = "F5", StyleIndex = (UInt32Value)2U };
            Cell cell223 = new Cell() { CellReference = "G5", StyleIndex = (UInt32Value)3U };
            Cell cell224 = new Cell() { CellReference = "H5", StyleIndex = (UInt32Value)3U };

            row28.Append(cell217);
            row28.Append(cell218);
            row28.Append(cell219);
            row28.Append(cell220);
            row28.Append(cell221);
            row28.Append(cell222);
            row28.Append(cell223);
            row28.Append(cell224);

            Row row29 = new Row() { RowIndex = (UInt32Value)6U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, Height = 18D, DyDescent = 0.25D };

            Cell cell225 = new Cell() { CellReference = "A6", StyleIndex = (UInt32Value)5U };
            CellValue cellValue76 = new CellValue();
            cellValue76.Text = "44166.661615277779";

            cell225.Append(cellValue76);
            Cell cell226 = new Cell() { CellReference = "B6", StyleIndex = (UInt32Value)4U };
            Cell cell227 = new Cell() { CellReference = "C6", StyleIndex = (UInt32Value)4U };
            Cell cell228 = new Cell() { CellReference = "D6", StyleIndex = (UInt32Value)1U };
            Cell cell229 = new Cell() { CellReference = "E6", StyleIndex = (UInt32Value)2U };
            Cell cell230 = new Cell() { CellReference = "F6", StyleIndex = (UInt32Value)2U };
            Cell cell231 = new Cell() { CellReference = "G6", StyleIndex = (UInt32Value)3U };
            Cell cell232 = new Cell() { CellReference = "H6", StyleIndex = (UInt32Value)3U };

            row29.Append(cell225);
            row29.Append(cell226);
            row29.Append(cell227);
            row29.Append(cell228);
            row29.Append(cell229);
            row29.Append(cell230);
            row29.Append(cell231);
            row29.Append(cell232);

            Row row30 = new Row() { RowIndex = (UInt32Value)7U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell233 = new Cell() { CellReference = "A7", StyleIndex = (UInt32Value)6U, DataType = CellValues.SharedString };
            CellValue cellValue77 = new CellValue();
            cellValue77.Text = "3";

            cell233.Append(cellValue77);
            Cell cell234 = new Cell() { CellReference = "B7", StyleIndex = (UInt32Value)7U };
            Cell cell235 = new Cell() { CellReference = "C7", StyleIndex = (UInt32Value)8U };
            Cell cell236 = new Cell() { CellReference = "D7", StyleIndex = (UInt32Value)8U };
            Cell cell237 = new Cell() { CellReference = "E7", StyleIndex = (UInt32Value)9U };
            Cell cell238 = new Cell() { CellReference = "F7", StyleIndex = (UInt32Value)9U };
            Cell cell239 = new Cell() { CellReference = "G7", StyleIndex = (UInt32Value)8U };
            Cell cell240 = new Cell() { CellReference = "H7", StyleIndex = (UInt32Value)8U };

            row30.Append(cell233);
            row30.Append(cell234);
            row30.Append(cell235);
            row30.Append(cell236);
            row30.Append(cell237);
            row30.Append(cell238);
            row30.Append(cell239);
            row30.Append(cell240);

            Row row31 = new Row() { RowIndex = (UInt32Value)8U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell241 = new Cell() { CellReference = "A8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue78 = new CellValue();
            cellValue78.Text = "5";

            cell241.Append(cellValue78);

            Cell cell242 = new Cell() { CellReference = "B8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue79 = new CellValue();
            cellValue79.Text = "6";

            cell242.Append(cellValue79);

            Cell cell243 = new Cell() { CellReference = "C8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue80 = new CellValue();
            cellValue80.Text = "7";

            cell243.Append(cellValue80);

            Cell cell244 = new Cell() { CellReference = "D8", StyleIndex = (UInt32Value)18U, DataType = CellValues.SharedString };
            CellValue cellValue81 = new CellValue();
            cellValue81.Text = "8";

            cell244.Append(cellValue81);

            Cell cell245 = new Cell() { CellReference = "E8", StyleIndex = (UInt32Value)20U, DataType = CellValues.SharedString };
            CellValue cellValue82 = new CellValue();
            cellValue82.Text = "9";

            cell245.Append(cellValue82);

            Cell cell246 = new Cell() { CellReference = "F8", StyleIndex = (UInt32Value)20U, DataType = CellValues.SharedString };
            CellValue cellValue83 = new CellValue();
            cellValue83.Text = "10";

            cell246.Append(cellValue83);

            Cell cell247 = new Cell() { CellReference = "G8", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue84 = new CellValue();
            cellValue84.Text = "11";

            cell247.Append(cellValue84);

            Cell cell248 = new Cell() { CellReference = "H8", StyleIndex = (UInt32Value)13U, DataType = CellValues.SharedString };
            CellValue cellValue85 = new CellValue();
            cellValue85.Text = "12";

            cell248.Append(cellValue85);

            row31.Append(cell241);
            row31.Append(cell242);
            row31.Append(cell243);
            row31.Append(cell244);
            row31.Append(cell245);
            row31.Append(cell246);
            row31.Append(cell247);
            row31.Append(cell248);

            Row row32 = new Row() { RowIndex = (UInt32Value)9U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };
            Cell cell249 = new Cell() { CellReference = "A9", StyleIndex = (UInt32Value)19U };
            Cell cell250 = new Cell() { CellReference = "B9", StyleIndex = (UInt32Value)19U };
            Cell cell251 = new Cell() { CellReference = "C9", StyleIndex = (UInt32Value)19U };
            Cell cell252 = new Cell() { CellReference = "D9", StyleIndex = (UInt32Value)19U };
            Cell cell253 = new Cell() { CellReference = "E9", StyleIndex = (UInt32Value)21U };
            Cell cell254 = new Cell() { CellReference = "F9", StyleIndex = (UInt32Value)21U };
            Cell cell255 = new Cell() { CellReference = "G9", StyleIndex = (UInt32Value)14U };
            Cell cell256 = new Cell() { CellReference = "H9", StyleIndex = (UInt32Value)14U };

            row32.Append(cell249);
            row32.Append(cell250);
            row32.Append(cell251);
            row32.Append(cell252);
            row32.Append(cell253);
            row32.Append(cell254);
            row32.Append(cell255);
            row32.Append(cell256);

            Row row33 = new Row() { RowIndex = (UInt32Value)10U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell257 = new Cell() { CellReference = "A10", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue86 = new CellValue();
            cellValue86.Text = "13";

            cell257.Append(cellValue86);

            Cell cell258 = new Cell() { CellReference = "B10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue87 = new CellValue();
            cellValue87.Text = "7812.89";

            cell258.Append(cellValue87);

            Cell cell259 = new Cell() { CellReference = "C10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue88 = new CellValue();
            cellValue88.Text = "7823.63";

            cell259.Append(cellValue88);

            Cell cell260 = new Cell() { CellReference = "D10", StyleIndex = (UInt32Value)10U };
            CellValue cellValue89 = new CellValue();
            cellValue89.Text = "1827";

            cell260.Append(cellValue89);

            Cell cell261 = new Cell() { CellReference = "E10", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue90 = new CellValue();
            cellValue90.Text = "14";

            cell261.Append(cellValue90);

            Cell cell262 = new Cell() { CellReference = "F10", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue91 = new CellValue();
            cellValue91.Text = "15";

            cell262.Append(cellValue91);

            Cell cell263 = new Cell() { CellReference = "G10", StyleIndex = (UInt32Value)12U };
            CellValue cellValue92 = new CellValue();
            cellValue92.Text = "1";

            cell263.Append(cellValue92);

            Cell cell264 = new Cell() { CellReference = "H10", StyleIndex = (UInt32Value)12U };
            CellValue cellValue93 = new CellValue();
            cellValue93.Text = "0.2";

            cell264.Append(cellValue93);

            row33.Append(cell257);
            row33.Append(cell258);
            row33.Append(cell259);
            row33.Append(cell260);
            row33.Append(cell261);
            row33.Append(cell262);
            row33.Append(cell263);
            row33.Append(cell264);

            Row row34 = new Row() { RowIndex = (UInt32Value)11U, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

            Cell cell265 = new Cell() { CellReference = "A11", StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
            CellValue cellValue94 = new CellValue();
            cellValue94.Text = "16";

            cell265.Append(cellValue94);

            Cell cell266 = new Cell() { CellReference = "B11", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula19 = new CellFormula();
            cellFormula19.Text = "SUM(B10:B10)";
            CellValue cellValue95 = new CellValue();
            cellValue95.Text = "7812.89";

            cell266.Append(cellFormula19);
            cell266.Append(cellValue95);

            Cell cell267 = new Cell() { CellReference = "C11", StyleIndex = (UInt32Value)10U };
            CellFormula cellFormula20 = new CellFormula();
            cellFormula20.Text = "SUM(C10:C10)";
            CellValue cellValue96 = new CellValue();
            cellValue96.Text = "7823.63";

            cell267.Append(cellFormula20);
            cell267.Append(cellValue96);

            Cell cell268 = new Cell() { CellReference = "D11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue97 = new CellValue();
            cellValue97.Text = "17";

            cell268.Append(cellValue97);

            Cell cell269 = new Cell() { CellReference = "E11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue98 = new CellValue();
            cellValue98.Text = "17";

            cell269.Append(cellValue98);

            Cell cell270 = new Cell() { CellReference = "F11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue99 = new CellValue();
            cellValue99.Text = "17";

            cell270.Append(cellValue99);

            Cell cell271 = new Cell() { CellReference = "G11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue100 = new CellValue();
            cellValue100.Text = "17";

            cell271.Append(cellValue100);

            Cell cell272 = new Cell() { CellReference = "H11", StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
            CellValue cellValue101 = new CellValue();
            cellValue101.Text = "17";

            cell272.Append(cellValue101);

            row34.Append(cell265);
            row34.Append(cell266);
            row34.Append(cell267);
            row34.Append(cell268);
            row34.Append(cell269);
            row34.Append(cell270);
            row34.Append(cell271);
            row34.Append(cell272);

            sheetData3.Append(row24);
            sheetData3.Append(row25);
            sheetData3.Append(row26);
            sheetData3.Append(row27);
            sheetData3.Append(row28);
            sheetData3.Append(row29);
            sheetData3.Append(row30);
            sheetData3.Append(row31);
            sheetData3.Append(row32);
            sheetData3.Append(row33);
            sheetData3.Append(row34);

            MergeCells mergeCells3 = new MergeCells() { Count = (UInt32Value)11U };
            MergeCell mergeCell23 = new MergeCell() { Reference = "H8:H9" };
            MergeCell mergeCell24 = new MergeCell() { Reference = "A1:H1" };
            MergeCell mergeCell25 = new MergeCell() { Reference = "A2:H2" };
            MergeCell mergeCell26 = new MergeCell() { Reference = "A3:H3" };
            MergeCell mergeCell27 = new MergeCell() { Reference = "A8:A9" };
            MergeCell mergeCell28 = new MergeCell() { Reference = "B8:B9" };
            MergeCell mergeCell29 = new MergeCell() { Reference = "C8:C9" };
            MergeCell mergeCell30 = new MergeCell() { Reference = "D8:D9" };
            MergeCell mergeCell31 = new MergeCell() { Reference = "E8:E9" };
            MergeCell mergeCell32 = new MergeCell() { Reference = "F8:F9" };
            MergeCell mergeCell33 = new MergeCell() { Reference = "G8:G9" };

            mergeCells3.Append(mergeCell23);
            mergeCells3.Append(mergeCell24);
            mergeCells3.Append(mergeCell25);
            mergeCells3.Append(mergeCell26);
            mergeCells3.Append(mergeCell27);
            mergeCells3.Append(mergeCell28);
            mergeCells3.Append(mergeCell29);
            mergeCells3.Append(mergeCell30);
            mergeCells3.Append(mergeCell31);
            mergeCells3.Append(mergeCell32);
            mergeCells3.Append(mergeCell33);
            PageMargins pageMargins3 = new PageMargins() { Left = 0.511811024D, Right = 0.511811024D, Top = 0.78740157499999996D, Bottom = 0.78740157499999996D, Header = 0.31496062000000002D, Footer = 0.31496062000000002D };

            worksheet3.Append(sheetDimension3);
            worksheet3.Append(sheetViews3);
            worksheet3.Append(sheetFormatProperties3);
            worksheet3.Append(columns3);
            worksheet3.Append(sheetData3);
            worksheet3.Append(mergeCells3);
            worksheet3.Append(pageMargins3);

            worksheetPart3.Worksheet = worksheet3;
        }

        // Generates content of sharedStringTablePart1.
        private void GenerateSharedStringTablePart1Content(SharedStringTablePart sharedStringTablePart1)
        {
            SharedStringTable sharedStringTable1 = new SharedStringTable() { Count = (UInt32Value)68U, UniqueCount = (UInt32Value)21U };

            SharedStringItem sharedStringItem1 = new SharedStringItem();
            Text text1 = new Text();
            text1.Text = "Relatório de Apoio à Conciliação - Parte 1";

            sharedStringItem1.Append(text1);

            SharedStringItem sharedStringItem2 = new SharedStringItem();
            Text text2 = new Text();
            text2.Text = "APLIC AUT MAIS - Posição em Ser Consolidada";

            sharedStringItem2.Append(text2);

            SharedStringItem sharedStringItem3 = new SharedStringItem();
            Text text3 = new Text();
            text3.Text = "Data Base: 30/11/2020";

            sharedStringItem3.Append(text3);

            SharedStringItem sharedStringItem4 = new SharedStringItem();
            Text text4 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text4.Text = " ";

            sharedStringItem4.Append(text4);

            SharedStringItem sharedStringItem5 = new SharedStringItem();
            Text text5 = new Text();
            text5.Text = "SPT/gpc";

            sharedStringItem5.Append(text5);

            SharedStringItem sharedStringItem6 = new SharedStringItem();
            Text text6 = new Text();
            text6.Text = "Cod. Ativo";

            sharedStringItem6.Append(text6);

            SharedStringItem sharedStringItem7 = new SharedStringItem();
            Text text7 = new Text();
            text7.Text = "Valor Aplicado";

            sharedStringItem7.Append(text7);

            SharedStringItem sharedStringItem8 = new SharedStringItem();
            Text text8 = new Text();
            text8.Text = "Valor Presente";

            sharedStringItem8.Append(text8);

            SharedStringItem sharedStringItem9 = new SharedStringItem();
            Text text9 = new Text();
            text9.Text = "Prazo";

            sharedStringItem9.Append(text9);

            SharedStringItem sharedStringItem10 = new SharedStringItem();
            Text text10 = new Text();
            text10.Text = "Emissão";

            sharedStringItem10.Append(text10);

            SharedStringItem sharedStringItem11 = new SharedStringItem();
            Text text11 = new Text();
            text11.Text = "Vencto";

            sharedStringItem11.Append(text11);

            SharedStringItem sharedStringItem12 = new SharedStringItem();
            Text text12 = new Text();
            text12.Text = "Tx. Ano / % CDI";

            sharedStringItem12.Append(text12);

            SharedStringItem sharedStringItem13 = new SharedStringItem();
            Text text13 = new Text();
            text13.Text = "Alíquota\n% Vigente";

            sharedStringItem13.Append(text13);

            SharedStringItem sharedStringItem14 = new SharedStringItem();
            Text text14 = new Text();
            text14.Text = "abcd";

            sharedStringItem14.Append(text14);

            SharedStringItem sharedStringItem15 = new SharedStringItem();
            Text text15 = new Text();
            text15.Text = "19/08/2020";

            sharedStringItem15.Append(text15);

            SharedStringItem sharedStringItem16 = new SharedStringItem();
            Text text16 = new Text();
            text16.Text = "20/08/2025";

            sharedStringItem16.Append(text16);

            SharedStringItem sharedStringItem17 = new SharedStringItem();
            Text text17 = new Text();
            text17.Text = "T O T A L";

            sharedStringItem17.Append(text17);

            SharedStringItem sharedStringItem18 = new SharedStringItem();
            Text text18 = new Text();
            text18.Text = "-";

            sharedStringItem18.Append(text18);

            SharedStringItem sharedStringItem19 = new SharedStringItem();
            Text text19 = new Text();
            text19.Text = "19/09/2020";

            sharedStringItem19.Append(text19);

            SharedStringItem sharedStringItem20 = new SharedStringItem();
            Text text20 = new Text();
            text20.Text = "20/08/2026";

            sharedStringItem20.Append(text20);

            SharedStringItem sharedStringItem21 = new SharedStringItem();
            Text text21 = new Text();
            text21.Text = "fghi";

            sharedStringItem21.Append(text21);

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

            sharedStringTablePart1.SharedStringTable = sharedStringTable1;
        }

        // Generates content of workbookStylesPart1.
        private void GenerateWorkbookStylesPart1Content(WorkbookStylesPart workbookStylesPart1)
        {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts1 = new Fonts() { Count = (UInt32Value)8U, KnownFonts = true };

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
            FontSize fontSize2 = new FontSize() { Val = 10D };
            FontName fontName2 = new FontName() { Val = "Arial" };

            font2.Append(fontSize2);
            font2.Append(fontName2);

            Font font3 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 16D };
            FontName fontName3 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };

            font3.Append(bold1);
            font3.Append(fontSize3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering2);

            Font font4 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize4 = new FontSize() { Val = 12D };
            FontName fontName4 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };

            font4.Append(bold2);
            font4.Append(fontSize4);
            font4.Append(fontName4);
            font4.Append(fontFamilyNumbering3);

            Font font5 = new Font();
            FontSize fontSize5 = new FontSize() { Val = 14D };
            FontName fontName5 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering4 = new FontFamilyNumbering() { Val = 2 };

            font5.Append(fontSize5);
            font5.Append(fontName5);
            font5.Append(fontFamilyNumbering4);

            Font font6 = new Font();
            Bold bold3 = new Bold();
            FontSize fontSize6 = new FontSize() { Val = 8D };
            FontName fontName6 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering5 = new FontFamilyNumbering() { Val = 2 };

            font6.Append(bold3);
            font6.Append(fontSize6);
            font6.Append(fontName6);
            font6.Append(fontFamilyNumbering5);

            Font font7 = new Font();
            FontSize fontSize7 = new FontSize() { Val = 10D };
            FontName fontName7 = new FontName() { Val = "Arial" };
            FontFamilyNumbering fontFamilyNumbering6 = new FontFamilyNumbering() { Val = 2 };

            font7.Append(fontSize7);
            font7.Append(fontName7);
            font7.Append(fontFamilyNumbering6);

            Font font8 = new Font();
            FontSize fontSize8 = new FontSize() { Val = 8D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName8 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering7 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font8.Append(fontSize8);
            font8.Append(color2);
            font8.Append(fontName8);
            font8.Append(fontFamilyNumbering7);
            font8.Append(fontScheme2);

            fonts1.Append(font1);
            fonts1.Append(font2);
            fonts1.Append(font3);
            fonts1.Append(font4);
            fonts1.Append(font5);
            fonts1.Append(font6);
            fonts1.Append(font7);
            fonts1.Append(font8);

            Fills fills1 = new Fills() { Count = (UInt32Value)2U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            fills1.Append(fill1);
            fills1.Append(fill2);

            Borders borders1 = new Borders() { Count = (UInt32Value)3U };

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
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color3);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color4);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color5);
            BottomBorder bottomBorder2 = new BottomBorder();
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            Border border3 = new Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder3.Append(color6);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color7 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder3.Append(color7);
            TopBorder topBorder3 = new TopBorder();

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder3.Append(color8);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            borders1.Append(border1);
            borders1.Append(border2);
            borders1.Append(border3);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = (UInt32Value)2U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats1.Append(cellFormat1);
            cellStyleFormats1.Append(cellFormat2);

            CellFormats cellFormats1 = new CellFormats() { Count = (UInt32Value)22U };
            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat4.Append(alignment1);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)3U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat5.Append(alignment2);
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)4U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true };

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat7.Append(alignment3);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)22U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Left, Vertical = VerticalAlignmentValues.Center };

            cellFormat8.Append(alignment4);

            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat9.Append(alignment5);

            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat10.Append(alignment6);
            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true };
            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)6U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true };
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFont = true };
            CellFormat cellFormat14 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, QuotePrefix = true, ApplyFont = true };
            CellFormat cellFormat15 = new CellFormat() { NumberFormatId = (UInt32Value)9U, FontId = (UInt32Value)7U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyNumberFormat = true, ApplyFont = true };

            CellFormat cellFormat16 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment7 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat16.Append(alignment7);
            CellFormat cellFormat17 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)1U, ApplyBorder = true };

            CellFormat cellFormat18 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyAlignment = true };
            Alignment alignment8 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true };

            cellFormat18.Append(alignment8);
            CellFormat cellFormat19 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyAlignment = true };

            CellFormat cellFormat20 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment9 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Justify };

            cellFormat20.Append(alignment9);

            CellFormat cellFormat21 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment10 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat21.Append(alignment10);

            CellFormat cellFormat22 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)1U, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment11 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat22.Append(alignment11);

            CellFormat cellFormat23 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)5U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment12 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center };

            cellFormat23.Append(alignment12);

            CellFormat cellFormat24 = new CellFormat() { NumberFormatId = (UInt32Value)49U, FontId = (UInt32Value)1U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)1U, ApplyNumberFormat = true, ApplyBorder = true, ApplyAlignment = true };
            Alignment alignment13 = new Alignment() { Vertical = VerticalAlignmentValues.Center };

            cellFormat24.Append(alignment13);

            cellFormats1.Append(cellFormat3);
            cellFormats1.Append(cellFormat4);
            cellFormats1.Append(cellFormat5);
            cellFormats1.Append(cellFormat6);
            cellFormats1.Append(cellFormat7);
            cellFormats1.Append(cellFormat8);
            cellFormats1.Append(cellFormat9);
            cellFormats1.Append(cellFormat10);
            cellFormats1.Append(cellFormat11);
            cellFormats1.Append(cellFormat12);
            cellFormats1.Append(cellFormat13);
            cellFormats1.Append(cellFormat14);
            cellFormats1.Append(cellFormat15);
            cellFormats1.Append(cellFormat16);
            cellFormats1.Append(cellFormat17);
            cellFormats1.Append(cellFormat18);
            cellFormats1.Append(cellFormat19);
            cellFormats1.Append(cellFormat20);
            cellFormats1.Append(cellFormat21);
            cellFormats1.Append(cellFormat22);
            cellFormats1.Append(cellFormat23);
            cellFormats1.Append(cellFormat24);

            CellStyles cellStyles1 = new CellStyles() { Count = (UInt32Value)2U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Normal", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };
            CellStyle cellStyle2 = new CellStyle() { Name = "Normal 2", FormatId = (UInt32Value)1U };

            cellStyles1.Append(cellStyle1);
            cellStyles1.Append(cellStyle2);
            DifferentialFormats differentialFormats1 = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            StylesheetExtensionList stylesheetExtensionList1 = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension1 = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension1.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles1 = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension1.Append(slicerStyles1);

            stylesheetExtensionList1.Append(stylesheetExtension1);

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

        // Generates content of themePart1.
        private void GenerateThemePart1Content(ThemePart themePart1)
        {
            A.Theme theme1 = new A.Theme() { Name = "Tema do Office" };
            theme1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements1 = new A.ThemeElements();

            A.ColorScheme colorScheme1 = new A.ColorScheme() { Name = "Escritório" };

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

            A.FontScheme fontScheme3 = new A.FontScheme() { Name = "Escritório" };

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
            A.SupplementalFont supplementalFont31 = new A.SupplementalFont() { Script = "Jpan", Typeface = "ＭＳ Ｐゴシック" };
            A.SupplementalFont supplementalFont32 = new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" };
            A.SupplementalFont supplementalFont33 = new A.SupplementalFont() { Script = "Hans", Typeface = "宋体" };
            A.SupplementalFont supplementalFont34 = new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" };
            A.SupplementalFont supplementalFont35 = new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" };
            A.SupplementalFont supplementalFont36 = new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" };
            A.SupplementalFont supplementalFont37 = new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" };
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

            fontScheme3.Append(majorFont1);
            fontScheme3.Append(minorFont1);

            A.FormatScheme formatScheme1 = new A.FormatScheme() { Name = "Escritório" };

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
            themeElements1.Append(fontScheme3);
            themeElements1.Append(formatScheme1);
            A.ObjectDefaults objectDefaults1 = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList1 = new A.ExtraColorSchemeList();

            theme1.Append(themeElements1);
            theme1.Append(objectDefaults1);
            theme1.Append(extraColorSchemeList1);

            themePart1.Theme = theme1;
        }

        private void SetPackageProperties(OpenXmlPackage document)
        {
            document.PackageProperties.Creator = "Jose Horacio Dias De Oliveira";
            document.PackageProperties.Created = System.Xml.XmlConvert.ToDateTime("2021-02-01T16:50:52Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.Modified = System.Xml.XmlConvert.ToDateTime("2021-02-02T11:07:05Z", System.Xml.XmlDateTimeSerializationMode.RoundtripKind);
            document.PackageProperties.LastModifiedBy = "Jose Horacio Dias De Oliveira";
        }


    }
}
