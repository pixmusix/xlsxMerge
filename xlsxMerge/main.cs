using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.IO;

namespace xlsxMerge
{
    public partial class main : Form
    {
        //Define xl Globals
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;

        Excel.Application xlOut;
        Excel.Workbook xlOutBook;

        public main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            InitSheetSelect(false);
            InitColumnSelect(false);
            InitRowSelect(false);

            xlOut = new Excel.Application();
            xlOut.Application.Workbooks.Add();
            xlOutBook = xlOut.Workbooks[1];
            xlOutBook.Worksheets.Add();
        }

        private void InitSheetSelect(Boolean b)
        {
            gbSheets.Visible = b;
            gbSheets.Enabled = b;

            if (xlWorkBook != null)
            {
                foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                {
                    cbLeftSheet.Items.Add(sheet.Name);
                    cbRightSheet.Items.Add(sheet.Name);
                }
            }
        }

        private void InitColumnSelect(Boolean b)
        {
            gbColumn.Visible = b;
            gbColumn.Enabled = b;

            if (xlWorkBook != null && b)
            {
                Excel.Worksheet xlWorkSheetLeft = xlWorkBook.Sheets[cbLeftSheet.SelectedItem.ToString()];
                numLeftKey.Maximum = xlWorkSheetLeft.UsedRange.Columns.Count;
                Excel.Worksheet xlWorkSheetRight = xlWorkBook.Sheets[cbRightSheet.SelectedItem.ToString()];
                numRightKey.Maximum = xlWorkSheetRight.UsedRange.Columns.Count;
            }
        }

        private void InitRowSelect(Boolean b)
        {
            gbRow.Visible = b;
            gbRow.Enabled = b;

            if (xlWorkBook != null && b)
            {
                Excel.Worksheet xlWorkSheetLeft = xlWorkBook.Sheets[cbLeftSheet.SelectedItem.ToString()];
                numLeftRow.Maximum = xlWorkSheetLeft.UsedRange.Rows.Count;
                Excel.Worksheet xlWorkSheetRight = xlWorkBook.Sheets[cbRightSheet.SelectedItem.ToString()];
                numRightRow.Maximum = xlWorkSheetRight.UsedRange.Rows.Count;
            }
        }

        private Excel.Worksheet Merge(Excel.Workbook book)
        {
            Excel.Worksheet king = xlWorkBook.Sheets[cbLeftSheet.SelectedItem.ToString()];
            Excel.Worksheet queen = xlWorkBook.Sheets[cbRightSheet.SelectedItem.ToString()];
            Excel.Worksheet jack = FullOuterJoin(king, queen);
            return jack;
        }

        private Excel.Worksheet FullOuterJoin(Excel.Worksheet A, Excel.Worksheet B)
        {
            int AKey = Convert.ToInt32(numLeftKey.Value);
            int BKey = Convert.ToInt32(numRightKey.Value);
            int ASta = Convert.ToInt32(numLeftRow.Value);
            int BSta = Convert.ToInt32(numRightRow.Value);

            List<int> AMat = new List<int>();
            List<int> BMat = new List<int>();

            Excel.Worksheet C = xlOutBook.Worksheets[1];
            C.Cells.ClearContents();
            int CSta = 1;

            for (int aj = ASta; aj < YRan(A) + 1; aj++)
            {
                for (int bj = BSta; bj < YRan(B) + 1; bj++)
                {
                    if (GetCell(A, aj, AKey) == GetCell(B, bj, BKey))
                    {
                        for (int i = 1; i < XRan(A) + 1; i++)
                        {
                            String cell = GetCell(A, aj, i);
                            C.Cells[CSta, i].Value = cell;
                        }
                        for (int i = 1; i < XRan(B) + 1; i++)
                        {
                            String cell = GetCell(B, bj, i);
                            C.Cells[CSta, i + XRan(A)].Value = cell;
                        }
                        CSta++;
                        AMat.Add(aj);
                        BMat.Add(bj);
                        break;
                    }
                }
            }

            for (int aj = ASta; aj < YRan(A) + 1; aj++)
            {
                if (AMat.Contains(aj)) { continue; }
                for (int i = 1; i < XRan(A) + 1; i++)
                {
                    String cell = GetCell(A, aj, i);
                    C.Cells[CSta, i].Value = cell;
                }
                CSta++;
            }
            for (int bj = BSta; bj < YRan(B) + 1; bj++)
            {
                if (BMat.Contains(bj)) { continue; }
                for (int i = 1; i < XRan(B) + 1; i++)
                {
                    String cell = GetCell(B, bj, i);
                    C.Cells[CSta, i + XRan(A)].Value = cell;
                }
                CSta++;
            }
            return C;
        }

        private String GetCell(Excel.Worksheet sheet, int x, int y, String ifnull = "__PRESERVE :: CELL_IS_EMPTY")
        {
            var cell = (sheet.Cells[x, y] as Excel.Range).Value;
            try { if (cell.ToString() == "__PRESERVE :: CELL_IS_EMPTY") { return ""; } } catch { }
            if (cell == null) { return ifnull; }
            return cell.ToString();
        }

        private int XRan(Excel.Worksheet sheet)
        {
            return sheet.UsedRange.Columns.Count;
        }

        private int YRan(Excel.Worksheet sheet)
        {
            return sheet.UsedRange.Rows.Count;
        }

        private DataTable ToDataTable(Excel.Worksheet sheet, int rDex)
        {
            DataTable df = new DataTable();
            for (int i = 0; i < XRan(sheet); i++)
            {
                df.Columns.Add(new DataColumn());
            }

            for (int j = rDex; j < YRan(sheet); j++)
            {
                DataRow df_row = df.NewRow();
                for (int i = 0; i < XRan(sheet); i++)
                {
                    String cell = GetCell(sheet, j + 1, i + 1, "Null");
                    df_row[i] = cell;
                }
                Console.WriteLine("<>");
                df.Rows.Add(df_row);
            }

            return df;
        }

        private void ReleaseObject(Object obj)
        {
            //Internet told me to do this.
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception) { obj = null; }
            finally { GC.Collect(); }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            String filePath = string.Empty;
            using (OpenFileDialog explorer = new OpenFileDialog())
            {
                explorer.InitialDirectory = "c:\\";
                explorer.Filter = "xlsx files (*.xlsx)|*.xlsx";
                explorer.FilterIndex = 2;
                explorer.RestoreDirectory = true;

                if (explorer.ShowDialog() == DialogResult.OK)
                {
                    filePath = explorer.FileName;
                }
            }

            Console.WriteLine(filePath);
            if (filePath.EndsWith(".xlsx"))
            {
                try
                {
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open Excel file \r\n >>\r\n" + ex.ToString());
                    return;
                }
                lblWorkbook.Text = xlWorkBook.Name;
            } 
            else
            {
                lblWorkbook.Text = "No Excel file Available";
                return;
            }
        }

        private void lblWorkbook_TextChanged(object sender, EventArgs e)
        {
            InitSheetSelect(true);
        }

        private void cbSheets_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cbLeftSheet.SelectedItem != null && cbRightSheet.SelectedItem != null)
            {
                InitColumnSelect(true);
                InitRowSelect(true);
                DataTable output = ToDataTable(Merge(xlWorkBook), 0);
                dgvOutput.DataSource = output;
            }
        }

        private void num_ValueChanged(object sender, EventArgs e)
        {
            DataTable dt = ToDataTable(Merge(xlWorkBook), 0);
            dgvOutput.DataSource = dt;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            DataTable dt = ToDataTable(Merge(xlWorkBook), 0);
            List<string> lines = new List<string>();
            EnumerableRowCollection<DataRow> edt = dt.AsEnumerable();
            EnumerableRowCollection<String> valueLines = edt.Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));
            lines.AddRange(valueLines);
            File.WriteAllLines(lblWorkbook.Text + "_MERGED.csv", lines);
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try { xlWorkBook.Close(); } catch (System.NullReferenceException) { }
            try { xlApp.Quit(); } catch (System.NullReferenceException) { }
            ReleaseObject(xlApp);
            ReleaseObject(xlWorkBook);
        }
    }
}
