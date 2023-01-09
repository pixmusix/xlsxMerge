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

        //Define Output Excel File
        Excel.Application xlOut;
        Excel.Workbook xlOutBook;

        public main()
        {
            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            // Init UI
            InitSheetSelect(false);
            InitColumnSelect(false);
            InitRowSelect(false);

            rbToCSV.Visible = false;
            rbToXSLX.Visible = false;
            rbToXSLX.Checked = true;
            btnSave.Enabled = false;

            //Prepare our output excel file
            xlOut = new Excel.Application();
            xlOut.Application.Workbooks.Add();
            xlOutBook = xlOut.Workbooks[1];
            xlOutBook.Worksheets.Add();
        }

        private void InitSheetSelect(Boolean b)
        {
            //Ensure load was successful
            if (xlWorkBook != null && b)
            {
                // Open up new Relevant UI for user
                gbSheets.Visible = b;
                gbSheets.Enabled = b;

                //Populate Combo Boxes
                cbLeftSheet.Items.Clear();
                cbRightSheet.Items.Clear();
                foreach (Excel.Worksheet sheet in xlWorkBook.Sheets)
                {
                    cbLeftSheet.Items.Add(sheet.Name);
                    cbRightSheet.Items.Add(sheet.Name);
                }
            } 
            else
            {
                gbSheets.Visible = b;
                gbSheets.Enabled = b;
                cbLeftSheet.Items.Clear();
                cbRightSheet.Items.Clear();
            }
        }

        private void InitColumnSelect(Boolean b)
        {
            // Ensure there is a workbook still in memory
            if (xlWorkBook != null && b)
            {
                // Open up new Relevant UI for user
                gbColumn.Visible = b;
                gbColumn.Enabled = b;

                //Set Number Box Maximums
                Excel.Worksheet xlWorkSheetLeft = xlWorkBook.Sheets[cbLeftSheet.SelectedItem.ToString()];
                numLeftKey.Maximum = xlWorkSheetLeft.UsedRange.Columns.Count;
                Excel.Worksheet xlWorkSheetRight = xlWorkBook.Sheets[cbRightSheet.SelectedItem.ToString()];
                numRightKey.Maximum = xlWorkSheetRight.UsedRange.Columns.Count;
            }
            else
            {
                gbColumn.Visible = b;
                gbColumn.Enabled = b;
            }
        }

        private void InitRowSelect(Boolean b)
        {
            // Ensure there is a workbook still in memory
            if (xlWorkBook != null && b) {

                // Open up new Relevant UI for user
                gbRow.Visible = b;
                gbRow.Enabled = b;

                //Set Number Box Maximums
                Excel.Worksheet xlWorkSheetLeft = xlWorkBook.Sheets[cbLeftSheet.SelectedItem.ToString()];
                numLeftRow.Maximum = xlWorkSheetLeft.UsedRange.Rows.Count;
                Excel.Worksheet xlWorkSheetRight = xlWorkBook.Sheets[cbRightSheet.SelectedItem.ToString()];
                numRightRow.Maximum = xlWorkSheetRight.UsedRange.Rows.Count;
            }
            else
            {
                gbRow.Visible = b;
                gbRow.Enabled = b;
            }
        }

        private Excel.Worksheet Merge(Excel.Workbook book)
        {
            // Get our worksheets from user input
            Excel.Worksheet king = xlWorkBook.Sheets[cbLeftSheet.SelectedItem.ToString()];
            Excel.Worksheet queen = xlWorkBook.Sheets[cbRightSheet.SelectedItem.ToString()];

            // **Merge!**
            Excel.Worksheet jack = FullOuterJoin(king, queen);
            return jack;
        }

        private Excel.Worksheet FullOuterJoin(Excel.Worksheet A, Excel.Worksheet B)
        {
            // Column Numbers (keys) and Starting row numbers from Users
            int AKey = Convert.ToInt32(numLeftKey.Value);
            int BKey = Convert.ToInt32(numRightKey.Value);
            int ASta = Convert.ToInt32(numLeftRow.Value);
            int BSta = Convert.ToInt32(numRightRow.Value);

            // A lookup table for matches so we don't double enter
            List<int> AMat = new List<int>();
            List<int> BMat = new List<int>();

            // Initilaise a new worksheet to populate
            Excel.Worksheet C = xlOutBook.Worksheets[1];
            C.Cells.ClearContents();
            int CSta = 1;

            // Check for B for a match with A using the col numbers provided by uesr.
            for (int aj = ASta; aj < YRan(A) + 1; aj++)
            {
                for (int bj = BSta; bj < YRan(B) + 1; bj++)
                {
                    if (GetCell(A, aj, AKey) == GetCell(B, bj, BKey))
                    {
                        // If we found a match, we populate the row with data from A & B
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
            // Fill rows from A with no neighbour into C
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
            // Fill rows from B with no neighbour into C
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
            // Get the cell
            var cell = (sheet.Cells[x, y] as Excel.Range).Value;
            // Preserve blank data (as opposed to empty merged nulls)
            try { if (cell.ToString() == "__PRESERVE :: CELL_IS_EMPTY") { return ""; } } catch { }
            // check for rows which had not match
            if (cell == null) { return ifnull; }
            // return the value
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
            //Initialise Columns
            DataTable df = new DataTable();
            for (int i = 0; i < XRan(sheet); i++)
            {
                df.Columns.Add(new DataColumn());
            }

            //Populate Rows
            for (int j = rDex; j < YRan(sheet); j++)
            {
                DataRow df_row = df.NewRow();
                for (int i = 0; i < XRan(sheet); i++)
                {
                    String cell = GetCell(sheet, j + 1, i + 1, "Null");
                    df_row[i] = cell;
                }
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

            //Open up a file explorer for user to retreive their xlsx file.
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

            //Attempt to convert file to workbook
            if (filePath.EndsWith(".xlsx"))
            {
                try
                {
                    //Populate the global xl variables
                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open Excel file \r\n >>\r\n" + ex.ToString());
                    return;
                }

                //User Feedback
                lblWorkbook.Text = xlWorkBook.Name;
            } 
            else
            {
                //User Feedback
                lblWorkbook.Text = "No Excel file Available";
                return;
            }
        }

        private void lblWorkbook_TextChanged(object sender, EventArgs e)
        {
            //(De)Initialise Relevant UI
            InitSheetSelect(true);
        }

        private void cbSheets_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cbLeftSheet.SelectedItem != null && cbRightSheet.SelectedItem != null)
            {
                //(De)Initialise Relevant UI
                InitColumnSelect(true);
                InitRowSelect(true);

                //Display data for user feedback
                DataTable dt = ToDataTable(Merge(xlWorkBook), 0);
                dgvOutput.DataSource = dt;
            }
        }

        private void num_ValueChanged(object sender, EventArgs e)
        {
            //Display Data for user feedback
            DataTable dt = ToDataTable(Merge(xlWorkBook), 0);
            dgvOutput.DataSource = dt;
        }

        private void rbToXSLX_CheckedChanged(object sender, EventArgs e)
        {
            rbToCSV.Checked = !rbToXSLX.Checked;
        }

        private void rbToCSV_CheckedChanged(object sender, EventArgs e)
        {
            rbToXSLX.Checked = !rbToCSV.Checked;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (rbToCSV.Checked) {
                DataTable dt = ToDataTable(Merge(xlWorkBook), 0);
                //Save to csv in local directory (Thanks internet).
                List<string> lines = new List<string>();
                EnumerableRowCollection<DataRow> edt = dt.AsEnumerable();
                EnumerableRowCollection<String> valueLines = edt.Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));
                lines.AddRange(valueLines);
                File.WriteAllLines(Environment.CurrentDirectory + "/" + lblWorkbook.Text + "_MERGED.csv", lines);
            }
            if (rbToXSLX.Checked)
            {
                Merge(xlWorkBook);
                //Sanitise
                for (int j = 1; j < YRan(xlOutBook.Worksheets[1]); j++)
                {
                    for (int i = 0; i < XRan(xlOutBook.Worksheets[1]); i++)
                    {
                        String cell = GetCell(xlOutBook.Worksheets[1], j + 1, i + 1, "Null");
                        xlOutBook.Worksheets[1].Cells[j + 1, i + 1] = cell;
                    }
                }
                try { xlOutBook.SaveAs(Environment.CurrentDirectory + "/" + lblWorkbook.Text + "_MERGED.xlsx"); } catch { }
            }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Release all of our excel spreadsheets from the Interop
            try { xlOutBook.Close(false, Type.Missing, Type.Missing); } catch (System.NullReferenceException) { }
            try { xlOut.Quit(); } catch (System.NullReferenceException) { }
            ReleaseObject(xlOutBook);
            ReleaseObject(xlOut);

            try { xlWorkBook.Close(false, Type.Missing, Type.Missing); } catch (System.NullReferenceException) { }
            try { xlApp.Quit(); } catch (System.NullReferenceException) { }
            ReleaseObject(xlWorkBook);
            ReleaseObject(xlApp);
        }

        private void dgvOutput_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            rbToCSV.Visible = dgvOutput.Rows.Count > 0;
            rbToXSLX.Visible = dgvOutput.Rows.Count > 0;
            btnSave.Enabled = dgvOutput.Rows.Count > 0;
        }

        private void dgvOutput_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            rbToCSV.Visible = dgvOutput.Rows.Count > 0;
            rbToXSLX.Visible = dgvOutput.Rows.Count > 0;
            btnSave.Enabled = dgvOutput.Rows.Count > 0;
        }
    }
}
