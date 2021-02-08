using System;
using System.IO;
using System.Windows.Forms;

namespace GetInfoFromWordToFireBirdTable
{
    public partial class MainForm : Form, IView
    {
        private readonly OpenFileDialog _openDBDialog;
        private readonly OpenFileDialog _openDocDialog;

        public event Action TableParseStarted;
        public event Action<string> CableNameChanged;
        public FileInfo FBDatabaseFile
        {
            get { return new FileInfo(dbFilePathTextBox.Text); }
        }

        public FileInfo MSWordFile
        {
            get { return new FileInfo(msWordFilePathTextBox.Text); }
        }

        public MainForm()
        {
            InitializeComponent();

            _openDBDialog = new OpenFileDialog()
            {
                Filter = "Firebird Database (*.fdb)|*.fdb",
                Title = "Выберите файл базы данных Firebird"
            };
            _openDocDialog = new OpenFileDialog()
            {
                Filter = "MS Word 2007 (*.docx)|*.docx|MS Word 2003 (*.doc)|*.doc",
                Title = "Выберите документ Microsoft Word"
            };

            openDBButton.Click += OpenDBButton_Click;
            openDocButton.Click += OpenDocButton_Click;
            cableBrandCheckComboBox.SelectedValueChanged += CableBrandCheckComboBox_SelectedValueChanged;
            wordTableParseStartButton.Click += WordTableParseStartButton_Click;

            progressBar.Step = 1;
        }

        public void SetCablesNames(string[] cablesNames)
        {
            cableBrandCheckComboBox.Items.AddRange(cablesNames);
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
            openDBButton.Enabled = true;
            openDocButton.Enabled = true;
            cableBrandCheckComboBox.Enabled = true;
        }

        private void CableBrandCheckComboBox_SelectedValueChanged(object sender, EventArgs e)
        {
            var cableName = cableBrandCheckComboBox.SelectedItem.ToString();
            CableNameChanged?.Invoke(cableName);
        }

        private void WordTableParseStartButton_Click(object sender, EventArgs e)
        {
            wordTableParseStartButton.Enabled = false;
            openDBButton.Enabled = false;
            openDocButton.Enabled = false;
            cableBrandCheckComboBox.Enabled = false;
            TableParseStarted?.Invoke();
        }

        private void OpenDocButton_Click(object sender, EventArgs e)
        {
            if (_openDocDialog.ShowDialog() == DialogResult.OK)
                msWordFilePathTextBox.Text = _openDocDialog.FileName;
        }

        private void OpenDBButton_Click(object sender, EventArgs e)
        {
            if (_openDBDialog.ShowDialog() == DialogResult.OK)
                dbFilePathTextBox.Text = _openDBDialog.FileName;
        }

    }
}
