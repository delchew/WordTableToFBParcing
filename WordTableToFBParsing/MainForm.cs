using GuiPresenter;
using System;
using System.IO;
using System.Windows.Forms;

namespace WordTableToFBParsing
{
    public partial class MainForm : Form, IView
    {
        private readonly OpenFileDialog _openDocDialog;

        public event Action TableParseStarted;
        public event Action<string> CableNameChanged;
        public event Action<string> DBConnectionNameChanged;

        public FileInfo MSWordFile
        {
            get { return new FileInfo(msWordFilePathTextBox.Text); }
        }

        public string DBConnectionName
        {
            get { return dbConnectionCheckComboBox.Text; }
        }

        public MainForm()
        {
            InitializeComponent();

            _openDocDialog = new OpenFileDialog()
            {
                Filter = "MS Word 2007 (*.docx)|*.docx|MS Word 2003 (*.doc)|*.doc",
                Title = "Выберите документ Microsoft Word"
            };

            openDocButton.Click += OpenDocButton_Click;
            cableBrandCheckComboBox.SelectedValueChanged += CableBrandCheckComboBox_SelectedValueChanged;
            dbConnectionCheckComboBox.SelectedValueChanged += DbConnectionCheckComboBox_SelectedValueChanged;
            wordTableParseStartButton.Click += WordTableParseStartButton_Click;

            progressBar.Step = 1;
        }

        public void SetCablesNames(string[] cablesNames)
        {
            cableBrandCheckComboBox.Items.AddRange(cablesNames);
        }

        public void SetDBConnectionsNames(string[] connectionsNames)
        {
            dbConnectionCheckComboBox.Items.AddRange(connectionsNames);
        }

        public void UpdateProgress(int parseOperationsCount, int completedOperationsCount)
        {
            Action action = () =>
            {
                if (progressBar.Value == 0)
                    progressBar.Maximum = parseOperationsCount;
                progressBar.Value = completedOperationsCount;
            };
            if (this.InvokeRequired)
                this.Invoke(action);
            else
                action();
        }

        public void ParseFinishReport()
        {
            wordTableParseStartButton.Enabled = true;
            dbConnectionCheckComboBox.Enabled = true;
            openDocButton.Enabled = true;
            cableBrandCheckComboBox.Enabled = true;
        }

        private void CableBrandCheckComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var cableName = cableBrandCheckComboBox.SelectedItem.ToString();
            CableNameChanged?.Invoke(cableName);
        }

        private void DbConnectionCheckComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var dbConnectionName = dbConnectionCheckComboBox.SelectedItem.ToString();
            DBConnectionNameChanged?.Invoke(dbConnectionName);
        }

        private void WordTableParseStartButton_Click(object sender, EventArgs e)
        {
            wordTableParseStartButton.Enabled = false;
            openDocButton.Enabled = false;
            cableBrandCheckComboBox.Enabled = false;
            dbConnectionCheckComboBox.Enabled = false;
            TableParseStarted?.Invoke();
        }

        private void OpenDocButton_Click(object sender, EventArgs e)
        {
            if (_openDocDialog.ShowDialog() == DialogResult.OK)
                msWordFilePathTextBox.Text = _openDocDialog.FileName;
        }
    }
}
