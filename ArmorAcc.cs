using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Globalization;




namespace ArmorAcc
{
    public partial class ArmorAcc
    {
        private void ArmorAcc_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnUpper_Click(object sender, RibbonControlEventArgs e)
        {
            Range mSelection = Globals.ThisAddIn.GetSelection();
            if (mSelection.Count >200)
            {
                MessageBox.Show("Vùng chọn quá lớn, Max 200 Cells!!","VƯỢT QUÁ 200 CELLS !!", MessageBoxButtons.OK);
                return;
            }

            foreach (Range mCell in mSelection)
            {
                if (mCell.Value != null && mCell.Value is string)
                {
                    mCell.Value = mCell.Value.ToUpper();
                }    
            }

        }

        private void btnLower_Click(object sender, RibbonControlEventArgs e)
        {
            Range mSelection = Globals.ThisAddIn.GetSelection();
            if (mSelection.Count > 200)
            {
                MessageBox.Show("Vùng chọn quá lớn, Max 200 Cells!!", "VƯỢT QUÁ 200 CELLS !!", MessageBoxButtons.OK);
                return;
            }

            foreach (Range mCell in mSelection)
            {
                if (mCell.Value != null && mCell.Value is string)
                {
                    mCell.Value = mCell.Value.ToLower();
                }
            }
        }

        private void btnProper_Click(object sender, RibbonControlEventArgs e)
        {
            Range mSelection = Globals.ThisAddIn.GetSelection();
            if (mSelection.Count > 200)
            {
                MessageBox.Show("Vùng chọn quá lớn, Max 200 Cells!!", "VƯỢT QUÁ 200 CELLS !!", MessageBoxButtons.OK);
                return;
            }

            foreach (Range mCell in mSelection)
            {
                TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
                if (mCell.Value != null && mCell.Value is string)
                {
                    mCell.Value = textInfo.ToTitleCase(mCell.Value.ToLower());
                }


            }
        }

        private void btnYelRow_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet mWs = Globals.ThisAddIn.GetActiveSheet();
            Range LastCell = Globals.ThisAddIn.GetLastCellFilled();
            Range mRow = mWs.Range[mWs.Cells[Globals.ThisAddIn.GetSelection().Row, 1], mWs.Cells[Globals.ThisAddIn.GetSelection().Row+ Globals.ThisAddIn.GetSelection().Rows.Count-1, LastCell.Column]];
            mRow.Interior.ColorIndex = 6;
        }

        private void btnRedRow_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet mWs = Globals.ThisAddIn.GetActiveSheet();
            Range LastCell = Globals.ThisAddIn.GetLastCellFilled();
            Range mRow = mWs.Range[mWs.Cells[Globals.ThisAddIn.GetSelection().Row, 1], mWs.Cells[Globals.ThisAddIn.GetSelection().Row + Globals.ThisAddIn.GetSelection().Rows.Count - 1, LastCell.Column]];
            mRow.Interior.ColorIndex = 3;

        }

        private void btnNofRow_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet mWs = Globals.ThisAddIn.GetActiveSheet();
            Range LastCell = Globals.ThisAddIn.GetLastCellFilled();
            Range mRow = mWs.Range[mWs.Cells[Globals.ThisAddIn.GetSelection().Row, 1], mWs.Cells[Globals.ThisAddIn.GetSelection().Row + Globals.ThisAddIn.GetSelection().Rows.Count - 1, LastCell.Column]];
            mRow.Interior.ColorIndex = 0;

        }

        private void btnTiltleoRow_Click(object sender, RibbonControlEventArgs e)
        {
            Range mSelection = Globals.ThisAddIn.GetSelection();
            Worksheet mWs = Globals.ThisAddIn.GetActiveSheet();
            Range LastCell = Globals.ThisAddIn.GetLastCellFilled();
            Range mRow = mWs.Range[mWs.Cells[mSelection.Row, 1], mWs.Cells[mSelection.Row + mSelection.Rows.Count - 1, LastCell.Column]];
            mRow.Interior.ColorIndex = 23;
            Font mFont = mRow.Font;
                mFont.Name = "Arial";
                mFont.Bold = true;
                mFont.Italic = false;
                mFont.ColorIndex = 2;
                mFont.Size = 14;
            if (mRow.Count > 200)
            {
                MessageBox.Show("Vùng chọn quá lớn, Max 200 Cells!!", "VƯỢT QUÁ 200 CELLS !!", MessageBoxButtons.OK);
                return;
            }

            foreach (Range mCell in mRow)
            {
                if (mCell.Value != null && mCell.Value is string)
                {
                    mCell.Value = mCell.Value.ToUpper();
                }
            }
        }

        private void btnSubRow_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet mWs = Globals.ThisAddIn.GetActiveSheet();
            Range LastCell = Globals.ThisAddIn.GetLastCellFilled();
            Range mRow = mWs.Range[mWs.Cells[Globals.ThisAddIn.GetSelection().Row, 1], mWs.Cells[Globals.ThisAddIn.GetSelection().Row + Globals.ThisAddIn.GetSelection().Rows.Count - 1, LastCell.Column]];
            mRow.Interior.ColorIndex = 6;
            Font mFont = mRow.Font;
            mFont.Name = "Arial";
            mFont.Bold = false;
            mFont.Italic = true;
            mFont.ColorIndex = 1;
            mFont.Size = 14;
        }

        private void btnTotalRow_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet mWs = Globals.ThisAddIn.GetActiveSheet();
            Range LastCell = Globals.ThisAddIn.GetLastCellFilled();
            Range mRow = mWs.Range[mWs.Cells[Globals.ThisAddIn.GetSelection().Row, 1], mWs.Cells[Globals.ThisAddIn.GetSelection().Row + Globals.ThisAddIn.GetSelection().Rows.Count - 1, LastCell.Column]];
            mRow.Interior.ColorIndex = 6;
            Font mFont = mRow.Font;
                mFont.Name = "Arial";
                mFont.Bold = true;
                mFont.Italic = false;
                mFont.ColorIndex = 3;
                mFont.Size = 14;
        }

        private void btnAutoF2_Click(object sender, RibbonControlEventArgs e)
        {
            Range mSelection = Globals.ThisAddIn.GetSelection();
            if (mSelection.Count > 2000)
            {
                MessageBox.Show("Vùng chọn quá lớn, Max 2.000 Cells!!", "VƯỢT QUÁ 2.000 CELLS !!", MessageBoxButtons.OK);
                return;
            }
            foreach (Range mCell in mSelection)
            {
                string mString = mCell.Value2 + "";
                mCell.Value = mString;
            }
        }

        private void btnOpnFileLocation_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook _ActiveWb = Globals.ThisAddIn.GetActiveWorkbook();

            string _myPath = _ActiveWb.FullName.Replace(_ActiveWb.Name, "");
            _myPath ='"' + _myPath + '"';
            System.Diagnostics.Process.Start("explorer.exe", @_myPath );

        }
        

    }
}
