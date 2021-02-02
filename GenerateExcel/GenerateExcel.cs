using DocumentFormat.OpenXml.Packaging;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using A = DocumentFormat.OpenXml.Drawing;

namespace ExcelOpenXML.GenerateExcel
{
    public partial class GenerateExcel
    {
        private void GenerateWorksheetContent(WorksheetPart worksheetPart3, int start)
        {
            Worksheet worksheet3 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet3.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet3.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet3.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            SheetDimension sheetDimension3 = new SheetDimension() { Reference = "A1:H65540" };

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

            sheetData3.Append(row24);
            sheetData3.Append(row25);
            sheetData3.Append(row26);
            sheetData3.Append(row27);
            sheetData3.Append(row28);
            sheetData3.Append(row29);
            sheetData3.Append(row30);
            sheetData3.Append(row31);
            sheetData3.Append(row32);

            //Linhas valor
            for (int i = 0; i < 65535; i++)
            {
                int linha = 10 + i;
                Row row33 = new Row() { RowIndex = (UInt32Value)(uint)linha, Spans = new ListValue<StringValue>() { InnerText = "1:8" }, DyDescent = 0.25D };

                Cell cell257 = new Cell() { CellReference = "A" + linha, StyleIndex = (UInt32Value)10U, DataType = CellValues.SharedString };
                CellValue cellValue86 = new CellValue();
                cellValue86.Text = "13";

                cell257.Append(cellValue86);

                Cell cell258 = new Cell() { CellReference = "B" + linha, StyleIndex = (UInt32Value)10U };
                CellValue cellValue87 = new CellValue();
                cellValue87.Text = "7812.89";

                cell258.Append(cellValue87);

                Cell cell259 = new Cell() { CellReference = "C" + linha, StyleIndex = (UInt32Value)10U };
                CellValue cellValue88 = new CellValue();
                cellValue88.Text = "7823.63";

                cell259.Append(cellValue88);

                Cell cell260 = new Cell() { CellReference = "D" + linha, StyleIndex = (UInt32Value)10U };
                CellValue cellValue89 = new CellValue();
                cellValue89.Text = "1827";

                cell260.Append(cellValue89);

                Cell cell261 = new Cell() { CellReference = "E" + linha, StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
                CellValue cellValue90 = new CellValue();
                cellValue90.Text = "14";

                cell261.Append(cellValue90);

                Cell cell262 = new Cell() { CellReference = "F" + linha, StyleIndex = (UInt32Value)11U, DataType = CellValues.SharedString };
                CellValue cellValue91 = new CellValue();
                cellValue91.Text = "15";

                cell262.Append(cellValue91);

                Cell cell263 = new Cell() { CellReference = "G" + linha, StyleIndex = (UInt32Value)12U };
                CellValue cellValue92 = new CellValue();
                cellValue92.Text = "1";

                cell263.Append(cellValue92);

                Cell cell264 = new Cell() { CellReference = "H" + linha, StyleIndex = (UInt32Value)12U };
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

                sheetData3.Append(row33);
            }

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

            //sheetData3.Append(row34);

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
    }
}
