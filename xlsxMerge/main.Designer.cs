
namespace xlsxMerge
{
    partial class main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblTitle = new System.Windows.Forms.Label();
            this.btnLoad = new System.Windows.Forms.Button();
            this.gbLoad = new System.Windows.Forms.GroupBox();
            this.lblWorkbook = new System.Windows.Forms.Label();
            this.gbSheets = new System.Windows.Forms.GroupBox();
            this.lblRightSheet = new System.Windows.Forms.Label();
            this.lblLeftSheet = new System.Windows.Forms.Label();
            this.cbRightSheet = new System.Windows.Forms.ComboBox();
            this.cbLeftSheet = new System.Windows.Forms.ComboBox();
            this.gbColumn = new System.Windows.Forms.GroupBox();
            this.numRightKey = new System.Windows.Forms.NumericUpDown();
            this.lblRightKey = new System.Windows.Forms.Label();
            this.numLeftKey = new System.Windows.Forms.NumericUpDown();
            this.lblLeftKey = new System.Windows.Forms.Label();
            this.gbRow = new System.Windows.Forms.GroupBox();
            this.numRightRow = new System.Windows.Forms.NumericUpDown();
            this.numLeftRow = new System.Windows.Forms.NumericUpDown();
            this.lblRightRowNum = new System.Windows.Forms.Label();
            this.lblLeftRowNum = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.dgvOutput = new System.Windows.Forms.DataGridView();
            this.gbLoad.SuspendLayout();
            this.gbSheets.SuspendLayout();
            this.gbColumn.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRightKey)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLeftKey)).BeginInit();
            this.gbRow.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRightRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLeftRow)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOutput)).BeginInit();
            this.SuspendLayout();
            // 
            // lblTitle
            // 
            this.lblTitle.AutoSize = true;
            this.lblTitle.Font = new System.Drawing.Font("Times New Roman", 24F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTitle.Location = new System.Drawing.Point(12, 9);
            this.lblTitle.Name = "lblTitle";
            this.lblTitle.Size = new System.Drawing.Size(152, 36);
            this.lblTitle.TabIndex = 1;
            this.lblTitle.Text = "xlsxMerge";
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(15, 19);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(86, 25);
            this.btnLoad.TabIndex = 4;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // gbLoad
            // 
            this.gbLoad.Controls.Add(this.lblWorkbook);
            this.gbLoad.Controls.Add(this.btnLoad);
            this.gbLoad.Location = new System.Drawing.Point(18, 68);
            this.gbLoad.Name = "gbLoad";
            this.gbLoad.Size = new System.Drawing.Size(305, 57);
            this.gbLoad.TabIndex = 5;
            this.gbLoad.TabStop = false;
            this.gbLoad.Text = "Load Workbook";
            // 
            // lblWorkbook
            // 
            this.lblWorkbook.AutoSize = true;
            this.lblWorkbook.Location = new System.Drawing.Point(117, 25);
            this.lblWorkbook.Name = "lblWorkbook";
            this.lblWorkbook.Size = new System.Drawing.Size(119, 13);
            this.lblWorkbook.TabIndex = 5;
            this.lblWorkbook.Text = "No Workbook Selected";
            this.lblWorkbook.TextChanged += new System.EventHandler(this.lblWorkbook_TextChanged);
            // 
            // gbSheets
            // 
            this.gbSheets.Controls.Add(this.lblRightSheet);
            this.gbSheets.Controls.Add(this.lblLeftSheet);
            this.gbSheets.Controls.Add(this.cbRightSheet);
            this.gbSheets.Controls.Add(this.cbLeftSheet);
            this.gbSheets.Location = new System.Drawing.Point(18, 131);
            this.gbSheets.Name = "gbSheets";
            this.gbSheets.Size = new System.Drawing.Size(305, 88);
            this.gbSheets.TabIndex = 6;
            this.gbSheets.TabStop = false;
            this.gbSheets.Text = "Sheet Select";
            // 
            // lblRightSheet
            // 
            this.lblRightSheet.AutoSize = true;
            this.lblRightSheet.Location = new System.Drawing.Point(159, 28);
            this.lblRightSheet.Name = "lblRightSheet";
            this.lblRightSheet.Size = new System.Drawing.Size(63, 13);
            this.lblRightSheet.TabIndex = 9;
            this.lblRightSheet.Text = "Right Sheet";
            // 
            // lblLeftSheet
            // 
            this.lblLeftSheet.AutoSize = true;
            this.lblLeftSheet.Location = new System.Drawing.Point(12, 28);
            this.lblLeftSheet.Name = "lblLeftSheet";
            this.lblLeftSheet.Size = new System.Drawing.Size(56, 13);
            this.lblLeftSheet.TabIndex = 6;
            this.lblLeftSheet.Text = "Left Sheet";
            // 
            // cbRightSheet
            // 
            this.cbRightSheet.FormattingEnabled = true;
            this.cbRightSheet.Location = new System.Drawing.Point(162, 53);
            this.cbRightSheet.Name = "cbRightSheet";
            this.cbRightSheet.Size = new System.Drawing.Size(137, 21);
            this.cbRightSheet.TabIndex = 8;
            this.cbRightSheet.SelectedValueChanged += new System.EventHandler(this.cbSheets_SelectedValueChanged);
            // 
            // cbLeftSheet
            // 
            this.cbLeftSheet.FormattingEnabled = true;
            this.cbLeftSheet.Location = new System.Drawing.Point(6, 53);
            this.cbLeftSheet.Name = "cbLeftSheet";
            this.cbLeftSheet.Size = new System.Drawing.Size(140, 21);
            this.cbLeftSheet.TabIndex = 7;
            this.cbLeftSheet.SelectedValueChanged += new System.EventHandler(this.cbSheets_SelectedValueChanged);
            // 
            // gbColumn
            // 
            this.gbColumn.Controls.Add(this.numRightKey);
            this.gbColumn.Controls.Add(this.lblRightKey);
            this.gbColumn.Controls.Add(this.numLeftKey);
            this.gbColumn.Controls.Add(this.lblLeftKey);
            this.gbColumn.Location = new System.Drawing.Point(18, 236);
            this.gbColumn.Name = "gbColumn";
            this.gbColumn.Size = new System.Drawing.Size(305, 57);
            this.gbColumn.TabIndex = 10;
            this.gbColumn.TabStop = false;
            this.gbColumn.Text = "Column Select";
            // 
            // numRightKey
            // 
            this.numRightKey.Location = new System.Drawing.Point(235, 25);
            this.numRightKey.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numRightKey.Name = "numRightKey";
            this.numRightKey.Size = new System.Drawing.Size(64, 20);
            this.numRightKey.TabIndex = 13;
            this.numRightKey.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numRightKey.ValueChanged += new System.EventHandler(this.num_ValueChanged);
            // 
            // lblRightKey
            // 
            this.lblRightKey.AutoSize = true;
            this.lblRightKey.Location = new System.Drawing.Point(159, 28);
            this.lblRightKey.Name = "lblRightKey";
            this.lblRightKey.Size = new System.Drawing.Size(53, 13);
            this.lblRightKey.TabIndex = 9;
            this.lblRightKey.Text = "Right Key";
            // 
            // numLeftKey
            // 
            this.numLeftKey.Location = new System.Drawing.Point(82, 25);
            this.numLeftKey.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numLeftKey.Name = "numLeftKey";
            this.numLeftKey.Size = new System.Drawing.Size(64, 20);
            this.numLeftKey.TabIndex = 12;
            this.numLeftKey.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numLeftKey.ValueChanged += new System.EventHandler(this.num_ValueChanged);
            // 
            // lblLeftKey
            // 
            this.lblLeftKey.AutoSize = true;
            this.lblLeftKey.Location = new System.Drawing.Point(12, 28);
            this.lblLeftKey.Name = "lblLeftKey";
            this.lblLeftKey.Size = new System.Drawing.Size(46, 13);
            this.lblLeftKey.TabIndex = 6;
            this.lblLeftKey.Text = "Left Key";
            // 
            // gbRow
            // 
            this.gbRow.Controls.Add(this.numRightRow);
            this.gbRow.Controls.Add(this.numLeftRow);
            this.gbRow.Controls.Add(this.lblRightRowNum);
            this.gbRow.Controls.Add(this.lblLeftRowNum);
            this.gbRow.Location = new System.Drawing.Point(18, 299);
            this.gbRow.Name = "gbRow";
            this.gbRow.Size = new System.Drawing.Size(305, 54);
            this.gbRow.TabIndex = 11;
            this.gbRow.TabStop = false;
            this.gbRow.Text = "Starting Row";
            // 
            // numRightRow
            // 
            this.numRightRow.Location = new System.Drawing.Point(235, 26);
            this.numRightRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numRightRow.Name = "numRightRow";
            this.numRightRow.Size = new System.Drawing.Size(64, 20);
            this.numRightRow.TabIndex = 11;
            this.numRightRow.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numRightRow.ValueChanged += new System.EventHandler(this.num_ValueChanged);
            // 
            // numLeftRow
            // 
            this.numLeftRow.Location = new System.Drawing.Point(82, 26);
            this.numLeftRow.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numLeftRow.Name = "numLeftRow";
            this.numLeftRow.Size = new System.Drawing.Size(64, 20);
            this.numLeftRow.TabIndex = 10;
            this.numLeftRow.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numLeftRow.ValueChanged += new System.EventHandler(this.num_ValueChanged);
            // 
            // lblRightRowNum
            // 
            this.lblRightRowNum.AutoSize = true;
            this.lblRightRowNum.Location = new System.Drawing.Point(159, 28);
            this.lblRightRowNum.Name = "lblRightRowNum";
            this.lblRightRowNum.Size = new System.Drawing.Size(67, 13);
            this.lblRightRowNum.TabIndex = 9;
            this.lblRightRowNum.Text = "Right Row #";
            // 
            // lblLeftRowNum
            // 
            this.lblLeftRowNum.AutoSize = true;
            this.lblLeftRowNum.Location = new System.Drawing.Point(12, 28);
            this.lblLeftRowNum.Name = "lblLeftRowNum";
            this.lblLeftRowNum.Size = new System.Drawing.Size(60, 13);
            this.lblLeftRowNum.TabIndex = 6;
            this.lblLeftRowNum.Text = "Left Row #";
            // 
            // btnSave
            // 
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSave.Location = new System.Drawing.Point(237, 382);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(86, 25);
            this.btnSave.TabIndex = 6;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // dgvOutput
            // 
            this.dgvOutput.AllowUserToAddRows = false;
            this.dgvOutput.AllowUserToDeleteRows = false;
            this.dgvOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvOutput.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvOutput.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvOutput.ColumnHeadersVisible = false;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvOutput.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgvOutput.Location = new System.Drawing.Point(342, 76);
            this.dgvOutput.Name = "dgvOutput";
            this.dgvOutput.ReadOnly = true;
            this.dgvOutput.RowHeadersVisible = false;
            this.dgvOutput.Size = new System.Drawing.Size(436, 331);
            this.dgvOutput.TabIndex = 12;
            // 
            // main
            // 
            this.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 430);
            this.Controls.Add(this.dgvOutput);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.gbRow);
            this.Controls.Add(this.gbColumn);
            this.Controls.Add(this.gbSheets);
            this.Controls.Add(this.gbLoad);
            this.Controls.Add(this.lblTitle);
            this.MinimumSize = new System.Drawing.Size(700, 450);
            this.Name = "main";
            this.Text = "3333";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.main_FormClosing);
            this.Load += new System.EventHandler(this.Main_Load);
            this.gbLoad.ResumeLayout(false);
            this.gbLoad.PerformLayout();
            this.gbSheets.ResumeLayout(false);
            this.gbSheets.PerformLayout();
            this.gbColumn.ResumeLayout(false);
            this.gbColumn.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRightKey)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLeftKey)).EndInit();
            this.gbRow.ResumeLayout(false);
            this.gbRow.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numRightRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numLeftRow)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgvOutput)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.GroupBox gbLoad;
        private System.Windows.Forms.Label lblWorkbook;
        private System.Windows.Forms.GroupBox gbSheets;
        private System.Windows.Forms.Label lblRightSheet;
        private System.Windows.Forms.Label lblLeftSheet;
        private System.Windows.Forms.ComboBox cbRightSheet;
        private System.Windows.Forms.ComboBox cbLeftSheet;
        private System.Windows.Forms.GroupBox gbColumn;
        private System.Windows.Forms.Label lblRightKey;
        private System.Windows.Forms.Label lblLeftKey;
        private System.Windows.Forms.GroupBox gbRow;
        private System.Windows.Forms.NumericUpDown numRightRow;
        private System.Windows.Forms.NumericUpDown numLeftRow;
        private System.Windows.Forms.Label lblRightRowNum;
        private System.Windows.Forms.Label lblLeftRowNum;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.DataGridView dgvOutput;
        private System.Windows.Forms.NumericUpDown numRightKey;
        private System.Windows.Forms.NumericUpDown numLeftKey;
    }
}

