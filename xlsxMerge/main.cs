using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
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

        //Define Primary Key Dictionaries
        Dictionary<String, Int32> Left_PrimaryKeys;
        Dictionary<String, Int32> Right_PrimaryKeys;
        List<Int32> Left_DouplicateKeys;
        List<Int32> Right_DouplicateKeys;

        //Define DataTable for user feedback
        DataTable dataframe;
        CancellationTokenSource cts;

        //Define UserInput Globals
        int dgvRowMax;
        int AIndex;
        int BIndex;
        int AFirstRow;
        int BFirstRow;
        String LeftSheetName;
        String RightSheetName;

        public main()
        {
            // Check if Excel is Installed on this machine and if not terminate.
            Type officeType = Type.GetTypeFromProgID("Excel.Application");
            if (officeType == null)
            {
                MessageBox.Show("Terminate : Excel Not Installed on this machine.");
                Environment.Exit(0);
            }

            InitializeComponent();
        }

        private void Main_Load(object sender, EventArgs e)
        {
            // Init UI
            InitSheetSelect(false);
            InitColumnSelect(false);
            InitRowSelect(false);

            lblLoading.Visible = false;
            rbToCSV.Visible = false;
            rbToXSLX.Visible = false;
            rbToXSLX.Checked = true;
            btnSave.Enabled = false;

            //Prepare our output excel file
            xlOut = new Excel.Application();
            xlOut.Application.Workbooks.Add();
            xlOutBook = xlOut.Workbooks[1];
            xlOutBook.Worksheets.Add();

            //Set the maximum rows to compute for dgvGridView User Feedback
            dgvRowMax = 100;
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

        private void InitOutput()
        {
            rbToCSV.Visible = dgvOutput.Rows.Count > 0;
            rbToXSLX.Visible = dgvOutput.Rows.Count > 0;
            btnSave.Enabled = dgvOutput.Rows.Count > 0;
        }

        private void FormatGrid()
        {
            if (dgvOutput.Columns.Count > 10)
            {
                dgvOutput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                dgvOutput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void Feedback()
        {
            lblLoading.Invoke((Action)(() => lblLoading.Visible = true));
            //Display Data for user feedback
            Excel.Worksheet sheet = xlOutBook.Worksheets.Add();
            if (!cts.IsCancellationRequested) { InitPrimaryKeys(); }
            if (!cts.IsCancellationRequested) { sheet = Merge(xlWorkBook); }
            if (!cts.IsCancellationRequested) { dataframe = ToDataTable(sheet, 0); }
            if (!cts.IsCancellationRequested)
            {
                dgvOutput.Invoke((Action)(() => dgvOutput.DataSource = dataframe));
                lblLoading.Invoke((Action)(() => lblLoading.Visible = false));
            } 
        }

        private void UpdateDataFrame()
        {
            //Get data from user input
            GetNumBoxes();

            // If there's a previous request, cancel it.
            if (cts != null)
            {
                cts.Cancel();
            }

            // Create a CTS for this request.
            cts = new CancellationTokenSource();

            // Update data grid view.
            try
            {
                Task rest_dgv = new Task(Feedback, cts.Token);
                rest_dgv.Start();
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("Operation Cancelled");
            }

        }

        private void GetNumBoxes()
        {
            AIndex = Convert.ToInt32(numLeftKey.Value);
            BIndex = Convert.ToInt32(numRightKey.Value);
            AFirstRow = Convert.ToInt32(numLeftRow.Value);
            BFirstRow = Convert.ToInt32(numRightRow.Value);
            LeftSheetName = cbLeftSheet.SelectedItem.ToString();
            RightSheetName = cbRightSheet.SelectedItem.ToString();
        }

        private void InitPrimaryKeys()
        {
            Left_PrimaryKeys = new Dictionary<String, Int32>();
            Right_PrimaryKeys = new Dictionary<String, Int32>();
            Left_DouplicateKeys = new List<Int32>();
            Right_DouplicateKeys = new List<Int32>();

            Excel.Worksheet left = xlWorkBook.Sheets[LeftSheetName];
            Excel.Worksheet right = xlWorkBook.Sheets[RightSheetName];

            for (int j = AFirstRow; j < YRan(left) + 1; j++)
            {
                try
                {
                    Left_PrimaryKeys.Add(GetCell(left, j, AIndex), j);
                }
                catch (System.ArgumentException)
                {
                    Left_DouplicateKeys.Add(j);
                }
            }
            for (int j = BFirstRow; j < YRan(right) + 1; j++)
            {
                try
                {
                    Right_PrimaryKeys.Add(GetCell(right, j, BIndex), j);
                }
                catch (System.ArgumentException)
                {
                    Right_DouplicateKeys.Add(j);
                }
            }
        }

        private Excel.Worksheet Merge(Excel.Workbook book)
        {
            // Get our worksheets from user input
            Excel.Worksheet king = xlWorkBook.Sheets[LeftSheetName];
            Excel.Worksheet queen = xlWorkBook.Sheets[RightSheetName];

            // **Merge!**
            Excel.Worksheet jack = FullOuterJoin(king, queen);
            return jack;
        }

        private Excel.Worksheet FullOuterJoin(Excel.Worksheet A, Excel.Worksheet B)
        {
            // Initilaise a new worksheet to populate
            Excel.Worksheet C = xlOutBook.Worksheets[1];
            C.Cells.ClearContents();
            int CSta = 1;

            if (Left_PrimaryKeys == null | Right_PrimaryKeys == null)
            {
                InitPrimaryKeys();
            }

            foreach (KeyValuePair<string, Int32> A_Entry in Left_PrimaryKeys)
            {
                int B_Row;
                if (Right_PrimaryKeys.TryGetValue(A_Entry.Key, out B_Row)) {
                    //If we found a match, we populate the row with data from A &B
                    for (int i = 1; i < XRan(A) + 1; i++)
                    {
                        String cell = GetCell(A, A_Entry.Value, i);
                        C.Cells[CSta, i].Value = cell;
                    }
                    for (int i = 1; i < XRan(B) + 1; i++)
                    {
                        String cell = GetCell(B, B_Row, i);
                        C.Cells[CSta, i + XRan(A)].Value = cell;
                    }
                    Right_PrimaryKeys.Remove(A_Entry.Key);
                }
                else
                {
                    for (int i = 1; i < XRan(A) + 1; i++)
                    {
                        String cell = GetCell(A, A_Entry.Value, i);
                        C.Cells[CSta, i].Value = cell;
                    }
                }
                CSta++;
            }


            // Fill rows from B with no neighbour into C
            foreach (KeyValuePair<string, Int32> B_Entry in Right_PrimaryKeys)
            {
                for (int i = 1; i < XRan(B) + 1; i++)
                {
                    String cell = GetCell(B, B_Entry.Value, i);
                    C.Cells[CSta, i + XRan(A)].Value = cell;
                }
                CSta++;
            }

            // Fill in Douplicate Keys
            foreach (Int32 row in Left_DouplicateKeys)
            {
                for (int i = 1; i < XRan(A) + 1; i++)
                {
                    String cell = GetCell(A, row, i);
                    C.Cells[CSta, i].Value = cell;
                }
                CSta++;
            }
            foreach (Int32 row in Right_DouplicateKeys)
            {
                for (int i = 1; i < XRan(B) + 1; i++)
                {
                    String cell = GetCell(B, row, i);
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
            // check for rows which had not match
            if (cell == null) { return ifnull; }
            // return the value
            String cellValue = cell.ToString();
            if (cellValue == "__PRESERVE :: CELL_IS_EMPTY") { return ""; }
            return cellValue;
        }

        private int XRan(Excel.Worksheet sheet)
        {
            return sheet.UsedRange.Columns.Count;
        }

        private int YRan(Excel.Worksheet sheet)
        {
            if (dgvRowMax < sheet.UsedRange.Rows.Count)
            {
                return dgvRowMax;
            }
            else
            {
                return sheet.UsedRange.Rows.Count;
            }
            
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

            ////Sort by keys
            //String sortColumn = df.Columns[AIndex].ColumnName + " ASC," + df.Columns[BIndex].ColumnName + " ASC";
            //df.DefaultView.Sort = sortColumn;
            //df = df.DefaultView.ToTable();
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
                UpdateDataFrame();
            }
        }


        private void num_ValueChanged(object sender, EventArgs e)
        {
            UpdateDataFrame();
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
            //We need to override the rowmax
            int cache = dgvRowMax;
            dgvRowMax = int.MaxValue;

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

            //restore the rowmax
            dgvRowMax = cache;
        }

        private void dgvOutput_DataSourceChanged(object sender, EventArgs e)
        {
            InitOutput();
            FormatGrid();
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
    }
}
