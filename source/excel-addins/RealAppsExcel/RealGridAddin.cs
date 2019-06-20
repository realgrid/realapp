using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace RealAppsExcel
{
    public partial class RealGridAddin
    {
        private CodeForm codeForm;

        private void RealGridAddin_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void RealGridAddin_Close(object sender, EventArgs e)
        {
            if (codeForm != null)
            {
                codeForm.Dispose();
                codeForm = null;
            }
        }

        private void BtnBuildForm_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet = app.ActiveSheet;
            String input = app.InputBox("컬럼 헤더의 세로 개수를 입력하세요", "컬럼 헤더 높이", 1);
            int depth = int.Parse(input);

            ColumnGenerator.BuildForm(sheet, depth);
        }

        private void BtnGenerateCode_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet = app.ActiveSheet;

            if (sheet.UsedRange.Columns.Count < 2 || sheet.UsedRange.Rows.Count < 2)
            {
                Utils.ShowMessage("시트 데이터가 존재하지 않습니다. 시트 서식화 후 진행하세요.");
            }
            else
            {
                string fieldText, columnText;
                string code = ColumnGenerator.GenerateCode(sheet, out fieldText, out columnText);
                if (!String.IsNullOrEmpty(code))
                {
                    if (this.codeForm == null)
                    {
                        codeForm = new CodeForm();
                    }
                    codeForm.FieldInfo = fieldText;
                    codeForm.ColumnInfo = columnText;
                    codeForm.SourceCode = code;
                    codeForm.ShowDialog();
                }
            }
        }

        private void BtrnExtractColumnWidth_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet = app.ActiveSheet;
            ColumnGenerator.ExtractWidth(sheet);     
        }

        private void BtnApplyColumnWidth_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet = app.ActiveSheet;
            ColumnGenerator.ApplyWidth(sheet);
        }

        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
             // cell width test
            Excel.Application app = Globals.ThisAddIn.Application;
            Excel.Worksheet sheet = app.ActiveSheet as Excel.Worksheet;
            Excel.Range defCell = app.ActiveCell;
            Utils.ShowMessage(app.ActiveWindow.PointsToScreenPixelsX(defCell.Width).ToString());

            double ratio = defCell.Width / defCell.ColumnWidth;
            Utils.ShowMessage("Width = " + defCell.Width.ToString() + ", ColumnWidth = " + defCell.ColumnWidth + ", Ratio = " + ratio.ToString());
            
            /* number validation test
            Excel.Range range = sheet.get_Range("A1", "A5") as Excel.Range;

            //delete previous validation rules 
            range.Validation.Delete();
            range.Validation.Add(Excel.XlDVType.xlValidateWholeNumber,
                                            Excel.XlDVAlertStyle.xlValidAlertStop,
                                            Excel.XlFormatConditionOperator.xlBetween,
                                            1, 1000);
            */
        }
    }
}
