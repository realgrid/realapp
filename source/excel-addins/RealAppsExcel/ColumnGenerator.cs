using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace RealAppsExcel
{
    class ColumnGenerator
    {

        // Property Count
        private const int PROP_COUNT = 9;
        // Property Names
        private static string HEAD_NAME = "필드명";
        private static string HEAD_TYPE = "형태";
        private static string HEAD_ALIGN = "정렬";
        private static string HEAD_EDIT = "편집 여부";
        private static string HEAD_NUMB = "숫자 형식";
        private static string HEAD_DATE = "날짜 형식";
        private static string HEAD_VALS = "선택 값";
        private static string HEAD_LBLS = "선택 라벨";
        private static string HEAD_SIZE = "너비";
        private static string[] PROP_TITLES = { HEAD_NAME, HEAD_TYPE, HEAD_ALIGN, HEAD_EDIT, HEAD_NUMB, HEAD_DATE, HEAD_VALS, HEAD_LBLS, HEAD_SIZE };

        // 기본 너비
        private static int DEF_WIDTH = 80;
        // 실제 데이터가 시작하는 행 번호
        private const int START_COL = 2;

        private static string[] TYPE_VALUES = { "line", "multiline", "dropdown", "search", "multicheck", "number", "date" };
        private static string[] TYPE_LABELS = { "텍스트", "다중 라인", "드랍다운", "부분검색", "복수 선택", "숫자", "날짜" };
        private static string[] ALIGN_VALUES = { "near", "center", "far" };
        private static string[] ALIGN_LABELS = { "왼쪽", "가운데", "오른쪽" };
        private static string[] EDIT_LABELS = { "편집 가능", "읽기 전용" };
        private static string[] NFMT_VALUES = { "", "#,##0", "#,##0.0", "#,##0.00" };
        private static string[] NFMT_LABELS = { "-", "정수", "소숫점 1자리", "소숫점 2자리" };
        private static string[] DFMT_VALUES = { "", "yyyy/MM/dd", "yyyy-MM-dd", "yyyy/MM/dd hh:nn:ss", "yyyy-MM-dd hh:nn:ss" };
        private static string[] DFMT_LABELS = { "-", "날짜('/')", "날짜('-')", "날짜('/') + 시간", "날짜('-') + 시간" };

        private static string ColName(int colIndex)
        {
            char col = Convert.ToChar(64 + colIndex);
            return col.ToString();
        }

        private static Range GetRange(Worksheet sheet, int row, int col, int rowCount = 1, int colCount = 1)
        {
            string str = ColName(col) + row;
            if (rowCount > 1 || colCount > 1)
            {
                str += ":" + ColName(col + colCount - 1) + (row + rowCount - 1);
            }
            return sheet.Range[str];
        }

        public static void BuildForm(Worksheet sheet, int colDepth)
        {
            int colCount = 4;
            int colPerGroup = 2;
            int XlSilverColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Silver);
            int XlBlackColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            int XlRedColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            int XlYellowColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
            
            sheet.UsedRange.Clear();
            sheet.UsedRange.ClearComments();
            sheet.UsedRange.Validation.Delete();
            
            // column group
            if (colDepth > 1)
            {
                colCount = Convert.ToInt32(Math.Pow(colPerGroup, colDepth));
                for (int r = 1; r < colDepth; r++)
                {
                    for (int c = 0; c < colCount; c++)
                    {
                        int groupCols = Convert.ToInt32(Math.Pow(colPerGroup, r));
                        if (c % groupCols == 0)
                        {
                            GetRange(sheet, colDepth - r, c + 2).Value = "그룹";
                            GetRange(sheet, colDepth - r, c + 2, 1, groupCols).Merge();
                        }
                    }
                }

                colCount++;
            }

            Range headerCells = GetRange(sheet, 1, 2, colDepth, colCount);
            headerCells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            headerCells.VerticalAlignment = XlVAlign.xlVAlignCenter;
            headerCells.Interior.Color = XlSilverColor;
            headerCells.Font.Size = 10;
            Borders borders = headerCells.Borders;
            borders.Item[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeLeft].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeRight].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeTop].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeBottom].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlInsideHorizontal].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlInsideVertical].Color = XlBlackColor;

            // column
            Range stdCell = sheet.Range["B1"];
            double defSize = Utils.PixelToColumnWidth(DEF_WIDTH);
            for (var i = 0; i < colCount; i++)
            {
                GetRange(sheet, colDepth, 2 + i).Value = "컬럼" + (i + 1);
                Range nameCell = GetRange(sheet, colDepth + 1, 2 + i);
                nameCell.Value = "column" + i;
                nameCell.ColumnWidth = defSize;
            }

            // left header
            int row = 1 + colDepth;
            int fieldNameRow = row;
            int editTypeRow = row + 1;
            int alignmentRow = row + 2;
            int editableRow = row + 3;
            int numberFmtRow = row + 4;
            int dateFmtRow = row + 5;
            int valuesRow = row + 6;
            int labelsRow = row + 7;
            int sizeRow = row + 8;

            for (int p = 0; p < PROP_COUNT; p++)
            {
                Range r = GetRange(sheet, row + p, 1);
                r.Value = PROP_TITLES[p];
                r.Font.Bold = true;
                if (p < 2)
                    r.Font.Color = XlRedColor;
                else if (p == 6)
                    r.AddComment("드랍다운 또는 복수선택일 경우\n항목들의 값을 쉼표(,)로 나누어 입력\n예)A,B,C");
                else if (p == 7)
                    r.AddComment("드랍다운 또는 복수선택일 경우\n항목들의 라벨을 쉼표(,)로 나누어 입력\n미입력시 값이 표시");
                else if (p == 8)
                    r.AddComment("너비의 값이 있으면 입력된 숫자값을 사용, 없으면 시트상의 컬럼의 너비를 사용");
            }

            Range bodyCells = GetRange(sheet, row, 2, PROP_COUNT, colCount);
            bodyCells.Interior.Color = XlYellowColor;
            bodyCells.Font.Size = 10;
            borders = bodyCells.Borders;
            borders.Item[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeLeft].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeRight].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlEdgeBottom].Color = XlBlackColor;
            borders.Item[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlInsideHorizontal].Color = XlSilverColor;
            borders.Item[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            borders.Item[XlBordersIndex.xlInsideVertical].Color = XlSilverColor;

            Range typeCells = GetRange(sheet, editTypeRow, 2, 1, colCount);
            typeCells.Validation.Add(XlDVType.xlValidateList, Type.Missing, Type.Missing, String.Join(",", TYPE_LABELS));
            typeCells.Value = TYPE_LABELS[0];

            Range alignCells = GetRange(sheet, alignmentRow, 2, 1, colCount);
            alignCells.Validation.Add(XlDVType.xlValidateList, Type.Missing, Type.Missing, String.Join(",", ALIGN_LABELS));
            alignCells.Value = ALIGN_LABELS[1];

            Range editableCells = GetRange(sheet, editableRow, 2, 1, colCount);
            editableCells.Validation.Add(XlDVType.xlValidateList, Type.Missing, Type.Missing, String.Join(",", EDIT_LABELS));
            editableCells.Value = EDIT_LABELS[0];

            Range numberCells = GetRange(sheet, numberFmtRow, 2, 1, colCount);
            numberCells.Validation.Add(XlDVType.xlValidateList, Type.Missing, Type.Missing, String.Join(",", NFMT_LABELS));
            numberCells.Value = NFMT_LABELS[0];

            Range dateCells = GetRange(sheet, dateFmtRow, 2, 1, colCount);
            dateCells.Validation.Add(XlDVType.xlValidateList, Type.Missing, Type.Missing, String.Join(",", DFMT_LABELS));
            dateCells.Value = NFMT_LABELS[0];

            Range sizeCells = GetRange(sheet, sizeRow, 2, 1, colCount);
            sizeCells.Validation.Delete();
            sizeCells.Validation.Add(XlDVType.xlValidateWholeNumber,
                                            XlDVAlertStyle.xlValidAlertStop,
                                            XlFormatConditionOperator.xlBetween,
                                            10, 10000);
        }

        // 컬럼의 그룹을 생성: recursively call
        private static int BuildColumnGroups(GridGroup group, Worksheet sheet, int row, int lastRow, int startCol, int lastCol, List<GridDataColumn> columns)
        {
            int totwidth = 0;
            for (var i = startCol; i <= lastCol; i++)
            {
                Range cell = GetRange(sheet, row, i);
                int mergeCount = cell.MergeArea.Cells.Count;
                if (mergeCount == 1)
                { // single column group
                    var col = columns[i - START_COL];
                    group.columns.Add(col);
                    totwidth += col.width;
                }
                else
                {
                    GridGroup child = new GridGroup();
                    child.setHeader(cell.Value);
                    child.width = BuildColumnGroups(child, sheet, row + 1, lastRow, i, i + mergeCount, columns);
                    totwidth += group.width;
                    group.columns.Add(child);
                    i += mergeCount - 1;
                }
            }
            return totwidth;
        }

        private static int FindSizeRow(Range allRange)
        {
            for (int r = 1; r <= allRange.Rows.Count; r++)
            {
                var head = allRange.Item[r, 1].Value;

                if (head == HEAD_SIZE)
                {
                    return r;
                }
            }
            return -1;
        }

        public static string GenerateCode(Worksheet sheet, out string fieldText, out string columnText)
        {
            int srow = -1;
            int nameRow = -1, typeRow = -1, alignRow = -1, editableRow = -1, numberRow = -1, dateRow = -1, valuesRow = -1, labelsRow = -1, sizeRow = -1;
            Range allRange = sheet.UsedRange;
            for (int r = 1; r <= allRange.Rows.Count; r++)
            {
                var head = allRange.Item[r, 1].Value;
                if (head != null && srow == -1)
                {
                    srow = r;
                }
                if (head == HEAD_NAME)
                {
                    nameRow = r;
                }
                else if (head == HEAD_TYPE)
                {
                    typeRow = r;
                }
                else if (head == HEAD_ALIGN)
                {
                    alignRow = r;
                }
                else if (head == HEAD_EDIT)
                {
                    editableRow = r;
                }
                else if (head == HEAD_NUMB)
                {
                    numberRow = r;
                }
                else if (head == HEAD_DATE)
                {
                    dateRow = r;
                }
                else if (head == HEAD_VALS)
                {
                    valuesRow = r;
                }
                else if (head == HEAD_LBLS)
                {
                    labelsRow = r;
                }
                else if (head == HEAD_SIZE)
                {
                    sizeRow = r;
                }
            }
            fieldText = "";
            columnText = "";
            if (nameRow < 1)
            {
                Utils.ShowMessage(HEAD_NAME + "행을 찾지 못했습니다.");
                return null;
            }
            if (typeRow < 1)
            {
                Utils.ShowMessage(HEAD_TYPE + "행을 찾지 못했습니다.");
                return null;
            }
            if (sizeRow < 1)
            {
                Utils.ShowMessage(HEAD_SIZE + "행을 찾지 못했습니다.");
                return null;
            }

            List<GridField> fields = new List<GridField>();
            List<GridDataColumn> columns = new List<GridDataColumn>();

            for (int c = 2; c <= allRange.Columns.Count; c++)
            {
                Range nameCell = allRange.Item[nameRow, c];
                string fname = nameCell.Value;
                double w = nameCell.Width / 72 * 96;
                if (!String.IsNullOrEmpty(fname))
                {
                    GridField field = new GridField();
                    GridDataColumn column = new GridDataColumn();
                    field.fieldName = column.fieldName = column.name = fname;
                    Range sizeCell = allRange.Item[sizeRow, c];
                    column.width = sizeCell.Value == null ? Utils.WidthToPixel(sizeCell.Width) : (int)sizeCell.Value;
                    column.setHeader(allRange.Item[srow - 1, c].Value);

                    string typeText = allRange.Item[typeRow, c].Value;
                    int typeIndex = Array.IndexOf(TYPE_LABELS, typeText);
                    if (typeIndex == -1)
                    {
                        Utils.ShowMessage("존재하지 않은 형태(" + typeText + ") 입니다.");
                        return null;
                    }
                    string ftype = TYPE_VALUES[typeIndex];
                    column.setEditor(ftype);

                    if (ftype == "number" || ftype == "datetime")
                    {
                        field.dataType = ftype;
                    }
                    string alignText = allRange[alignRow, c].Value;
                    int alignIndex = Array.IndexOf(ALIGN_VALUES, alignText);
                    string align = ALIGN_VALUES[alignIndex > -1 ? alignIndex : 0];
                    if (align != "near")
                    {
                        column.setStyles(align);
                    }

                    if (allRange[editableRow, c].Value == EDIT_LABELS[1])
                    {
                        column.editable = false;
                    }

                    if (ftype == "date")
                    {
                        string fmt = allRange[dateRow, c].Value;
                        int formatIndex = Array.IndexOf(DFMT_LABELS, fmt);
                        string dateformat = formatIndex == -1 ? String.IsNullOrEmpty(fmt) ? "yyyy/MM/dd" : fmt : DFMT_VALUES[formatIndex];
                        field.datetimeFormat = dateformat;
                        column.editor.datetimeFormat = dateformat;
                    }

                    if (ftype == "number")
                    {
                        var fmt = allRange[numberRow, c].Value;
                        var formatIndex = Array.IndexOf(NFMT_LABELS, fmt);
                        var numberformat = formatIndex == -1 ? fmt || null : NFMT_VALUES[formatIndex];
                        column.setStyles(null, numberformat);
                        column.editor.editFormat = numberformat;
                        column.editor.textAlignment = "far";
                    }

                    if (ftype == "multicheck" || ftype == "dropdown")
                    {
                        column.lookupDisplay = true;
                        column.editor.showButtons = true;
                        string valText = allRange[valuesRow, c].Value;
                        string lblText = allRange[labelsRow, c].Value;
                        column.values = valText;
                        if (!String.IsNullOrEmpty(lblText))
                            column.labels = lblText;
                        if (ftype == "multicheck")
                            column.valueSeperator = ",";
                    }

                    fields.Add(field);
                    columns.Add(column);
                }
            }
            JsonSerializerSettings settings = new JsonSerializerSettings();
            settings.NullValueHandling = NullValueHandling.Ignore;
            fieldText = JsonConvert.SerializeObject(fields, Formatting.Indented, settings);
            if (srow > 1)
            {
                GridGroup groups = new GridGroup();
                BuildColumnGroups(groups, sheet, 1, srow, 2, columns.Count + 1, columns);
                columnText = JsonConvert.SerializeObject(groups.columns, Formatting.Indented, settings);
                //return { fields: fields, columns: groups};
            }
            else
            {
                columnText = JsonConvert.SerializeObject(columns, Formatting.Indented, settings);
                //return { fields: fields, columns: columns};
            }
            string code = "var divId = \"realgrid\";\r\n" +
                "dataProvider = new RealGridJS.LocalDataProvider();\r\n" +
                "var fields = " + fieldText + ";\r\n" +
                "dataProvider.setFields(fields);\r\n" +
                "\r\n" +
                "gridView = new RealGridJS.GridView(divId);\r\n" +
                "var columns = " + columnText + ";\r\n" +
                "gridView.setDataSource(dataProvider);\r\n" +
                "gridView.setColumns(columns);";
            return code;
        }

        public static void ExtractWidth(Worksheet sheet)
        {
            Range allRange = sheet.UsedRange;
            int sizeRow = FindSizeRow(allRange);
            for (int c = 2; c <= allRange.Columns.Count; c++)
            {
                Range cell = allRange.Item[sizeRow, c];
                int pixel = Utils.WidthToPixel(cell.Width);
                cell.Value = pixel;
            }
        }

        public static void ApplyWidth(Worksheet sheet)
        {
            Range allRange = sheet.UsedRange;
            int sizeRow = FindSizeRow(allRange);
            for (int c = 2; c <= allRange.Columns.Count; c++)
            {
                Range cell = allRange.Item[sizeRow, c];
                if (cell.Value != null)
                {
                    cell.ColumnWidth = Utils.PixelToColumnWidth((int)cell.Value);
                }
            }

        }
    }
}
