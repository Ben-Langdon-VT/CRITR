namespace CRITR
{
    partial class frmCRITR
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnSelectExcel = new Button();
            btnSelectImageFolder = new Button();
            lblExcelFileName = new Label();
            btnSelectOutputFolder = new Button();
            fileSystemWatcher1 = new FileSystemWatcher();
            lblImageFolder = new Label();
            lblOutputFolder = new Label();
            lblOutputFileName = new Label();
            btnGenerateOutput = new Button();
            btnCancel = new Button();
            lblLogFileName = new Label();
            txtOutputFile = new TextBox();
            btnClear = new Button();
            rdWord = new RadioButton();
            rdPpt = new RadioButton();
            panel1 = new Panel();
            lblOutputDocumentType = new Label();
            ((System.ComponentModel.ISupportInitialize)fileSystemWatcher1).BeginInit();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // btnSelectExcel
            // 
            btnSelectExcel.Location = new Point(338, 65);
            btnSelectExcel.Name = "btnSelectExcel";
            btnSelectExcel.Size = new Size(368, 58);
            btnSelectExcel.TabIndex = 0;
            btnSelectExcel.Text = "Select Excel File";
            btnSelectExcel.UseVisualStyleBackColor = true;
            btnSelectExcel.Click += btnSelectExcel_Click;
            // 
            // btnSelectImageFolder
            // 
            btnSelectImageFolder.Location = new Point(338, 129);
            btnSelectImageFolder.Name = "btnSelectImageFolder";
            btnSelectImageFolder.Size = new Size(368, 58);
            btnSelectImageFolder.TabIndex = 1;
            btnSelectImageFolder.Text = "Select Image Folder";
            btnSelectImageFolder.UseVisualStyleBackColor = true;
            btnSelectImageFolder.Click += btnSelectImageFolder_Click;
            // 
            // lblExcelFileName
            // 
            lblExcelFileName.AutoSize = true;
            lblExcelFileName.Location = new Point(743, 74);
            lblExcelFileName.Name = "lblExcelFileName";
            lblExcelFileName.Size = new Size(0, 41);
            lblExcelFileName.TabIndex = 2;
            // 
            // btnSelectOutputFolder
            // 
            btnSelectOutputFolder.Location = new Point(338, 193);
            btnSelectOutputFolder.Name = "btnSelectOutputFolder";
            btnSelectOutputFolder.Size = new Size(368, 58);
            btnSelectOutputFolder.TabIndex = 3;
            btnSelectOutputFolder.Text = "Select Output Folder";
            btnSelectOutputFolder.UseVisualStyleBackColor = true;
            btnSelectOutputFolder.Click += btnSelectOutputFolder_Click;
            // 
            // fileSystemWatcher1
            // 
            fileSystemWatcher1.EnableRaisingEvents = true;
            fileSystemWatcher1.SynchronizingObject = this;
            // 
            // lblImageFolder
            // 
            lblImageFolder.AutoSize = true;
            lblImageFolder.Location = new Point(743, 138);
            lblImageFolder.Name = "lblImageFolder";
            lblImageFolder.Size = new Size(0, 41);
            lblImageFolder.TabIndex = 5;
            // 
            // lblOutputFolder
            // 
            lblOutputFolder.AutoSize = true;
            lblOutputFolder.Location = new Point(743, 202);
            lblOutputFolder.Name = "lblOutputFolder";
            lblOutputFolder.Size = new Size(0, 41);
            lblOutputFolder.TabIndex = 6;
            // 
            // lblOutputFileName
            // 
            lblOutputFileName.AutoSize = true;
            lblOutputFileName.Location = new Point(451, 438);
            lblOutputFileName.Name = "lblOutputFileName";
            lblOutputFileName.Size = new Size(260, 41);
            lblOutputFileName.TabIndex = 7;
            // 
            // btnGenerateOutput
            // 
            btnGenerateOutput.Location = new Point(338, 584);
            btnGenerateOutput.Name = "btnGenerateOutput";
            btnGenerateOutput.Size = new Size(369, 58);
            btnGenerateOutput.TabIndex = 8;
            btnGenerateOutput.Text = "Generate Output";
            btnGenerateOutput.UseVisualStyleBackColor = true;
            btnGenerateOutput.Click += btnGenerateOutput_Click;
            // 
            // btnCancel
            // 
            btnCancel.Location = new Point(779, 584);
            btnCancel.Name = "btnCancel";
            btnCancel.Size = new Size(369, 58);
            btnCancel.TabIndex = 9;
            btnCancel.Text = "Cancel";
            btnCancel.UseVisualStyleBackColor = true;
            btnCancel.Click += btnCancel_Click;
            // 
            // lblLogFileName
            // 
            lblLogFileName.AutoSize = true;
            lblLogFileName.Location = new Point(338, 680);
            lblLogFileName.Name = "lblLogFileName";
            lblLogFileName.Size = new Size(0, 41);
            lblLogFileName.TabIndex = 10;
            // 
            // txtOutputFile
            // 
            txtOutputFile.Location = new Point(743, 438);
            txtOutputFile.Name = "txtOutputFile";
            txtOutputFile.Size = new Size(801, 47);
            txtOutputFile.TabIndex = 11;
            // 
            // btnClear
            // 
            btnClear.Location = new Point(1220, 584);
            btnClear.Name = "btnClear";
            btnClear.Size = new Size(369, 58);
            btnClear.TabIndex = 12;
            btnClear.Text = "Clear Entries";
            btnClear.UseVisualStyleBackColor = true;
            btnClear.Click += btnClear_Click;
            // 
            // rdWord
            // 
            rdWord.AutoSize = true;
            rdWord.Location = new Point(32, 31);
            rdWord.Name = "rdWord";
            rdWord.Size = new Size(128, 45);
            rdWord.TabIndex = 13;
            rdWord.TabStop = true;
            rdWord.Text = "Word";
            rdWord.UseVisualStyleBackColor = true;
            rdWord.CheckedChanged += radioButton1_CheckedChanged;
            // 
            // rdPpt
            // 
            rdPpt.AutoSize = true;
            rdPpt.Location = new Point(357, 31);
            rdPpt.Name = "rdPpt";
            rdPpt.Size = new Size(207, 45);
            rdPpt.TabIndex = 14;
            rdPpt.TabStop = true;
            rdPpt.Text = "Powerpoint";
            rdPpt.UseVisualStyleBackColor = true;
            rdPpt.CheckedChanged += radioButton2_CheckedChanged;
            // 
            // panel1
            // 
            panel1.Controls.Add(rdWord);
            panel1.Controls.Add(rdPpt);
            panel1.Location = new Point(743, 285);
            panel1.Name = "panel1";
            panel1.Size = new Size(630, 116);
            panel1.TabIndex = 15;
            // 
            // lblOutputDocumentType
            // 
            lblOutputDocumentType.AutoSize = true;
            lblOutputDocumentType.Location = new Point(375, 315);
            lblOutputDocumentType.Name = "lblOutputDocumentType";
            lblOutputDocumentType.Size = new Size(0, 41);
            lblOutputDocumentType.TabIndex = 16;
            // 
            // frmCRITR
            // 
            AutoScaleDimensions = new SizeF(17F, 41F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1936, 881);
            Controls.Add(lblOutputDocumentType);
            Controls.Add(panel1);
            Controls.Add(btnClear);
            Controls.Add(txtOutputFile);
            Controls.Add(lblLogFileName);
            Controls.Add(btnCancel);
            Controls.Add(btnGenerateOutput);
            Controls.Add(lblOutputFileName);
            Controls.Add(lblOutputFolder);
            Controls.Add(lblImageFolder);
            Controls.Add(btnSelectOutputFolder);
            Controls.Add(lblExcelFileName);
            Controls.Add(btnSelectImageFolder);
            Controls.Add(btnSelectExcel);
            Name = "frmCRITR";
            Text = "CRITR";
            ((System.ComponentModel.ISupportInitialize)fileSystemWatcher1).EndInit();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        //select excel file
        //select input image folder
        //select output folder
        //enter output file name
        //run button
        //view debug button

        private Button btnSelectExcel;
        private Button btnSelectImageFolder;
        //private OpenFileDialog openFileDialog1;
        //private SaveFileDialog saveFileDialog1;
        private Label lblExcelFileName;
        private Button btnSelectOutputFolder;
        private FileSystemWatcher fileSystemWatcher1;
        private Label lblOutputFileName;
        private Label lblOutputFolder;
        private Label lblImageFolder;
        private Button btnGenerateOutput;
        private Button btnCancel;
        private Label lblLogFileName;
        private TextBox txtOutputFile;
        private Button btnClear;

        //application state variables
        private string txtExcelFile;
        private string txtImageFolder;
        private string txtOutputFolder;
        private string OutputFileName;
        private string txtOutputFormat;
        private bool boolWord;
        private bool boolPpt;

        private void SetCancelButton(Button myCancelBtn)
        {
            this.CancelButton = btnCancel;
        }

        private void ClearAppStateVariables()
        {
            txtExcelFile = "";
            txtImageFolder = "";
            txtOutputFolder = "";
            OutputFileName = "";
            txtOutputFormat = "";
            boolWord = true;
            boolPpt = false;
        }

        private void InitializeLabels()
        {
            lblExcelFileName.Text = "{excel file name}";
            lblImageFolder.Text = "{input image folder}";
            lblOutputDocumentType.Text = "Output Document Type:";
            lblOutputFolder.Text = "{output folder}";
            lblOutputFileName.Text = "Output File Name:";
            lblLogFileName.Text = "{type document name here}";
            rdWord.Checked = true;
            rdPpt.Checked = false;
        }

        private RadioButton rdPpt;
        private RadioButton rdWord;
        private Panel panel1;
        private Label lblOutputDocumentType;
    }
}