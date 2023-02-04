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

        //Define DataTable and boolean for user feedback
        Data DataFrame;
        DataTable MasterDataTable;
        bool workerRequested;



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

            //Set the worker request flag
            workerRequested = false;

            //Initialise Data Objects
            DataFrame = new Data();
            MasterDataTable = new DataTable();
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
            // Open up new Relevant UI for user
            gbColumn.Visible = b;
            gbColumn.Enabled = b;

            //Set Number Box Maximums
            numLeftKey.Maximum = DataFrame.DataV.Columns.Count;
            numRightKey.Maximum = DataFrame.DataW.Columns.Count;
        }

        private void InitRowSelect(Boolean b)
        {
            // Open up new Relevant UI for user
            gbRow.Visible = b;
            gbRow.Enabled = b;

            //Set Number Box Maximums
            numLeftKey.Maximum = DataFrame.DataV.Columns.Count;
            numRightKey.Maximum = DataFrame.DataW.Columns.Count;
        }

        private void InitOutput()
        {
            rbToCSV.Visible = dgvOutput.Rows.Count > 0;
            rbToXSLX.Visible = dgvOutput.Rows.Count > 0;
            btnSave.Enabled = dgvOutput.Rows.Count > 0;
        }

        private void FormatGrid()
        {
            if (dgvOutput.Columns.Count > 6)
            {
                dgvOutput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                dgvOutput.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            }
        }

        private void PassUserInterface()
        {
            DataFrame.VIndex = Convert.ToInt32(numLeftKey.Value);
            DataFrame.WIndex = Convert.ToInt32(numRightKey.Value);
            DataFrame.VFirstRow = Convert.ToInt32(numLeftRow.Value);
            DataFrame.WFirstRow = Convert.ToInt32(numRightRow.Value);
        }

        private void StartWorker()
        {
            lblLoading.Text = "Loading Preview...";
            lblLoading.Visible = true;
            if (workerFeedback.IsBusy != true)
            {
                workerRequested = false;
                GetUserInterface();
                workerFeedback.RunWorkerAsync();
            }
            else
            {
                workerRequested = true;
            }
        }

        private void CancelWorker()
        {
            if (workerFeedback.WorkerSupportsCancellation == true)
            {
                workerFeedback.CancelAsync();
            }
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
                //Create a new datatable
                LeftSheetName = cbLeftSheet.SelectedItem.ToString();
                RightSheetName = cbRightSheet.SelectedItem.ToString();

                Excel.Worksheet king = xlWorkBook.Sheets[LeftSheetName];
                Excel.Worksheet queen = xlWorkBook.Sheets[RightSheetName];

                DataFrame = new Data(king, queen);


                //Initialise Relevant UI
                InitColumnSelect(true);
                InitRowSelect(true);

                //Get the data 
                PassUserInterface();

                //Display data for user feedback
                if (workerFeedback.IsBusy)
                {
                    CancelWorker();
                }
                StartWorker();
            } 
            else
            {
                InitColumnSelect(false);
                InitRowSelect(false);
            }
        }

        private void num_ValueChanged(object sender, EventArgs e)
        {
            if (workerFeedback.IsBusy) {
                CancelWorker();
            }
            StartWorker();
        }

        private void workerFeedback_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            PassUserInterface();
            if (worker.CancellationPending) { e.Cancel = true; } else { DataFrame.InitPrimaryKeys(); }
            if (worker.CancellationPending) { e.Cancel = true; } else { DataFrame.Merge(xlWorkBook); }
            if (worker.CancellationPending) { e.Cancel = true; } else { MasterDataTable = ToDataTable(xlOutBook.Worksheets[1], 0); }
            dgvOutput.Invoke((Action)(() => dgvOutput.DataSource = MasterDataTable));
        }

        private void workerFeedback_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                Console.WriteLine("Worker.Canceled!");
            }
            else if (e.Error != null)
            {
                Console.WriteLine("Worker.ERROR!" + e.Error.Message);
                lblLoading.Text = "Error: " + e.Error.Message;
            }
            else
            {
                Console.WriteLine("Worker Completed Successfully! :)");
                lblLoading.Visible = false;
            }
            if (workerRequested)
            {
                StartWorker();
            }
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
            this.Cursor = Cursors.WaitCursor;

            //We need to override the rowmax

            //Keep our globals up to date
            PassUserInterface();
            InitPrimaryKeys();

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

            this.Cursor = Cursors.Default;
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

    public class Data
    {
        //Define Tables
        public DataTable DataV;
        public DataTable DataW;

        //Define UserInput Globals
        public int VIndex;
        public int WIndex;
        public int VFirstRow;
        public int WFirstRow;

        //Define Primary Key Dictionaries
        Dictionary<String, Int32> Left_PrimaryKeys;
        Dictionary<String, Int32> Right_PrimaryKeys;
        List<Int32> Left_DouplicateKeys;
        List<Int32> Right_DouplicateKeys;

        public Data(Excel.Worksheet V, Excel.Worksheet W)
        {
            DataV = ToDataTable(V);
            DataW = ToDataTable(W);
        }

        public Data()
        {
            DataV = new DataTable();
            DataW = new DataTable();
        }

        private DataTable ToDataTable(Excel.Worksheet sheet)
        {
            //Initialise Columns
            DataTable df = new DataTable();
            for (int i = 0; i < XRan(sheet); i++)
            {
                df.Columns.Add(new DataColumn());
            }

            //Populate Rows
            for (int j = 0; j < YRan(sheet); j++)
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
            return sheet.UsedRange.Rows.Count;
        }

        private DataTable SortDataTable(DataTable df) 
        {
            ////Sort by keys
            String sortColumn = df.Columns[VIndex].ColumnName + " ASC," + df.Columns[WIndex].ColumnName + " ASC";
            df.DefaultView.Sort = sortColumn;
            df = df.DefaultView.ToTable();
            return df;
        }

        public void InitPrimaryKeys()
        {
            Left_PrimaryKeys = new Dictionary<String, Int32>();
            Right_PrimaryKeys = new Dictionary<String, Int32>();
            Left_DouplicateKeys = new List<Int32>();
            Right_DouplicateKeys = new List<Int32>();

            for (int j = VFirstRow; j < DataV.Rows.Count; j++)
            {
                try
                {
                    Left_PrimaryKeys.Add(DataV.Rows[j][VIndex].ToString(), j);
                }
                catch (System.ArgumentException)
                {
                    Left_DouplicateKeys.Add(j);
                }
            }
            for (int j = WFirstRow; j < DataW.Rows.Count; j++)
            {
                try
                {
                    Right_PrimaryKeys.Add(DataV.Rows[j][VIndex].ToString(), j);
                }
                catch (System.ArgumentException)
                {
                    Right_DouplicateKeys.Add(j);
                }
            }
        }

        private DataTable FullOuterJoin(DataTable V, DataTable W)
        {
            // Initilaise a new worksheet to populate
            DataTable Z = new DataTable();
            int CSta = 0;

            if (Left_PrimaryKeys == null | Right_PrimaryKeys == null)
            {
                InitPrimaryKeys();
            }

            foreach (KeyValuePair<string, Int32> V_Entry in Left_PrimaryKeys)
            {
                int W_Row;
                if (Right_PrimaryKeys.TryGetValue(V_Entry.Key, out W_Row))
                {
                    //If we found a match, we populate the row with data from A &B
                    for (int i = 0; i < V.Rows.Count; i++)
                    {
                        String cell = V.Rows[V_Entry.Value][i].ToString();
                        Z.Rows[CSta][i] = cell;
                    }
                    for (int i = 0; i < W.Rows.Count; i++)
                    {
                        String cell = W.Rows[W_Row][i].ToString();
                        Z.Rows[CSta][i + V.Rows.Count] = cell;
                    }
                    Right_PrimaryKeys.Remove(V_Entry.Key);
                }
                else
                {
                    for (int i = 0; i < V.Rows.Count; i++)
                    {
                        String cell = V.Rows[V_Entry.Value][i].ToString();
                        Z.Rows[CSta][i] = cell;
                    }
                }
                CSta++;
            }


            // Fill rows from B with no neighbour into C
            foreach (KeyValuePair<string, Int32> W_Entry in Right_PrimaryKeys)
            {
                for (int i = 0; i < W.Rows.Count; i++)
                {
                    String cell = W.Rows[W_Entry.Value][i].ToString();
                    Z.Rows[CSta][i + V.Rows.Count] = cell;
                }
                CSta++;
            }

            // Fill in Douplicate Keys
            foreach (Int32 row in Left_DouplicateKeys)
            {
                for (int i = 0; i < V.Rows.Count; i++)
                {
                    String cell = V.Rows[row][i].ToString();
                    Z.Rows[CSta][i] = cell;
                }
                CSta++;
            }
            foreach (Int32 row in Right_DouplicateKeys)
            {
                for (int i = 0; i < W.Rows.Count; i++)
                {
                    String cell = W.Rows[row][i].ToString();
                    Z.Rows[CSta][i + V.Rows.Count] = cell;
                }
                CSta++;
            }

            return Z;
        }

    }
}
