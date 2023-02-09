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

        //Define DataTable and boolean for user feedback
        Data DataFrame;
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
            //Initialise Data Object
            DataFrame = new Data();

            // Init UI
            InitSheetSelect(false);
            InitColumnSelect(false);
            InitRowSelect(false);

            lblLoading.Visible = false;
            rbToCSV.Visible = false;
            rbToXSLX.Visible = false;
            rbToXSLX.Checked = true;
            btnSave.Enabled = false;

            //Set the worker request flag
            workerRequested = false;
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
            numLeftKey.Maximum = DataFrame.DataV.Columns.Count + 1;
            numRightKey.Maximum = DataFrame.DataW.Columns.Count + 1;
        }

        private void InitRowSelect(Boolean b)
        {
            // Open up new Relevant UI for user
            gbRow.Visible = b;
            gbRow.Enabled = b;

            //Set Number Box Maximums
            numLeftRow.Maximum = DataFrame.DataV.Rows.Count + 1;
            numRightRow.Maximum = DataFrame.DataW.Rows.Count + 1;
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
            DataFrame.VIndex = Convert.ToInt32(numLeftKey.Value) - 1;
            DataFrame.WIndex = Convert.ToInt32(numRightKey.Value) - 1;
            DataFrame.VFirstRow = Convert.ToInt32(numLeftRow.Value) - 1;
            DataFrame.WFirstRow = Convert.ToInt32(numRightRow.Value) - 1;
        }

        private Excel.Workbook To_XLSX(Data df, Excel.Workbook wb)
        {

            for (int j = 0; j < df.DataZ.Rows.Count; j++)
            {
                for (int i = 0; i < df.DataZ.Columns.Count; i++)
                {
                    String cell = df.DataZ.Rows[j][i].ToString();
                    wb.Worksheets[1].Cells[j + 1, i + 1] = cell;
                }
            }
            return wb;
        }

        private void StartWorker()
        {
            lblLoading.Text = "Loading Preview...";
            lblLoading.Visible = true;
            if (workerFeedback.IsBusy != true)
            {
                workerRequested = false;
                PassUserInterface();
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
                    InitSheetSelect(true);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Could not open Excel file \r\n >>\r\n" + ex.ToString());
                    InitSheetSelect(false);
                    return;
                }

                //User Feedback
                lblWorkbook.Text = xlWorkBook.Name;
            } 
            else
            {
                //User Feedback
                lblWorkbook.Text = "No Excel file Available";
                InitSheetSelect(false);
                return;
            }
        }

        private void cbSheets_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cbLeftSheet.SelectedItem != null & cbRightSheet.SelectedItem != null)
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

        private void dgvOutput_DataSourceChanged(object sender, EventArgs e)
        {
            InitOutput();
            FormatGrid();
        }

        private void workerFeedback_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            if (worker.CancellationPending) { e.Cancel = true; } else { DataFrame.InitPrimaryKeys(); }
            if (worker.CancellationPending) { e.Cancel = true; } else { DataFrame.Merge(); }
            dgvOutput.Invoke((Action)(() => dgvOutput.DataSource = DataFrame.DataZ));
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

            //Keep our globals up to date
            PassUserInterface();
            DataFrame.InitPrimaryKeys();
            DataFrame.Merge();
            if (rbToCSV.Checked) {
                DataTable dt = DataFrame.DataZ;

                //Save to csv in local directory (Thanks internet).
                List<string> lines = new List<string>();
                EnumerableRowCollection<DataRow> edt = dt.AsEnumerable();
                EnumerableRowCollection<String> valueLines = edt.Select(row => string.Join(",", row.ItemArray.Select(val => $"\"{val}\"")));
                lines.AddRange(valueLines);
                File.WriteAllLines(Environment.CurrentDirectory + "/" + lblWorkbook.Text + "_MERGED.csv", lines);
            }
            if (rbToXSLX.Checked)
            {
                //Prepare our output excel file
                Excel.Application xlOut = new Excel.Application();
                Excel.Workbook xlOutWorkbook = xlOut.Application.Workbooks.Add();
                xlOutWorkbook = xlOut.Workbooks[1];
                xlOutWorkbook = To_XLSX(DataFrame, xlOutWorkbook);
                //Save the file
                try { xlOutWorkbook.SaveAs(Environment.CurrentDirectory + "/" + lblWorkbook.Text + "_MERGED.xlsx"); } catch { }
                //Release all of our excel spreadsheets from the Interop
                try { xlOutWorkbook.Close(false, Type.Missing, Type.Missing); } catch (System.NullReferenceException) { }
                try { xlOut.Quit(); } catch (System.NullReferenceException) { }
                ReleaseObject(xlOutWorkbook);
                ReleaseObject(xlOut);
            }

            this.Cursor = Cursors.Default;
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Release all of our excel spreadsheets from the Interop
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
        public DataTable DataZ;

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
                    Right_PrimaryKeys.Add(DataW.Rows[j][WIndex].ToString(), j);
                }
                catch (System.ArgumentException)
                {
                    Right_DouplicateKeys.Add(j);
                }
            }
        }

        public DataTable Merge()
        {
            return FullOuterJoin(DataV, DataW);
        }

        private DataTable FullOuterJoin(DataTable V, DataTable W)
        {
            //Initilaise a new worksheet to populate
            DataTable Z = new DataTable();
            for (int i = 1; i < DataV.Columns.Count + DataW.Columns.Count; i++)
            {
                DataColumn column = new DataColumn();
                column.DataType = System.Type.GetType("System.String");
                column.DefaultValue = "Null";
                column.Unique = false;
                Z.Columns.Add(column);
            }
            int CSta = 0;

            if (Left_PrimaryKeys == null | Right_PrimaryKeys == null)
            {
                InitPrimaryKeys();
            }

            foreach (KeyValuePair<string, Int32> V_Entry in Left_PrimaryKeys)
            {
                Z.Rows.Add();
                if (Right_PrimaryKeys.TryGetValue(V_Entry.Key, out int W_Row))
                {
                    //If we found a match, we populate the row with data from A &B
                    for (int i = 0; i < V.Columns.Count; i++)
                    {
                        String cell = V.Rows[V_Entry.Value][i].ToString();
                        Z.Rows[CSta][i] = cell;
                    }
                    for (int i = 0; i < W.Columns.Count; i++)
                    {
                        String cell = W.Rows[W_Row][i].ToString();
                        Z.Rows[CSta][i + V.Rows.Count] = cell;
                    }
                    Right_PrimaryKeys.Remove(V_Entry.Key);
                }
                else
                {
                    for (int i = 0; i < V.Columns.Count; i++)
                    {
                        String cell = V.Rows[V_Entry.Value][i].ToString();        
                        Z.Rows[CSta][i] = cell;
                    }
                }
                CSta++;
            }


            //Fill rows from B with no neighbour into C
            foreach (KeyValuePair<string, Int32> W_Entry in Right_PrimaryKeys)
            {
                Z.Rows.Add();
                for (int i = 0; i < W.Columns.Count; i++)
                {
                    String cell = W.Rows[W_Entry.Value][i].ToString();
                    Z.Rows[CSta][i + V.Rows.Count] = cell;
                }
                CSta++;
            }

            //Fill in Douplicate Keys
            foreach (Int32 row in Left_DouplicateKeys)
            {
                Z.Rows.Add();
                for (int i = 0; i < V.Columns.Count; i++)
                {
                    String cell = V.Rows[row][i].ToString();
                    Z.Rows[CSta][i] = cell;
                }
                CSta++;
            }
            foreach (Int32 row in Right_DouplicateKeys)
            {
                Z.Rows.Add();
                for (int i = 0; i < W.Columns.Count; i++)
                {
                    String cell = W.Rows[row][i].ToString();
                    Z.Rows[CSta][i + V.Rows.Count] = cell;
                }
                CSta++;
            }

            //Save a copy of Z as property and pass it back to the namespace.
            DataZ = Z;
            return Z;
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
                    String cell = GetCell(sheet, j + 1, i + 1);
                    df_row[i] = cell;
                }
                df.Rows.Add(df_row);
            }

            return df;
        }

        private String GetCell(Excel.Worksheet sheet, int x, int y, String ifnull = "")
        {
            // Get the cell
            var cell = (sheet.Cells[x, y] as Excel.Range).Value;
            // return the value
            if (cell == null) 
            { 
                return ifnull; 
            }
            else
            {
                return cell.ToString();
            }
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

    }
    
}
