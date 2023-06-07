using System.Runtime.CompilerServices;
using System.Windows.Forms;

namespace WinFormsApp1
{
    public partial class frmCRITR : Form
    {
        public frmCRITR()
        {
            InitializeComponent();
            ClearAppStateVariables();
            InitializeLabels();
        }

        private void btnSelectExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            // Add code here to configure the dialog and set filters
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string filePath = Path.GetDirectoryName(openFileDialog1.FileName);

                // set the output variable values
                lblExcelFileName.Text = filePath + '/' + openFileDialog1.FileName;   //for display
                txtExcelFile = filePath;            //internal value
            }

        }

        private void btnGenerateOutput_Click(object sender, EventArgs e)
        {
            lblLogFileName.Text = "Log File Available!";
            // check validity of inputs
            
            // then run the command line

        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void btnSelectImageFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;
                    // Add code here to use the selected path
                    lblImageFolder.Text = selectedPath;
                    txtImageFolder = selectedPath;
                }
            }
        }

        private void btnSelectOutputFolder_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            {
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    string selectedPath = folderBrowserDialog.SelectedPath;
                    // Add code here to use the selected path
                    lblOutputFolder.Text = selectedPath;
                    txtOutputFolder = selectedPath;
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            ClearAppStateVariables();
            InitializeLabels();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}