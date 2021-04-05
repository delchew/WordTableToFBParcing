using System;
using System.IO;
using System.Threading;
using Gtk;
using GuiPresenter;
using UI = Gtk.Builder.ObjectAttribute;

namespace WordTableToFBParsingGTK
{
    class MainParseWindow : Window, IView
    {
        [UI] private Entry filePathEntry = null;
        [UI] private Button openDocButton = null;
        [UI] private Button parseButton = null;
        [UI] private ComboBoxText connectionCombobox = null;
        [UI] private ComboBoxText cableTypeCombobox = null;
        [UI] private ProgressBar progressBar = null;
        private FileChooserDialog chooserDialog;

        public event System.Action TableParseStarted;
        public event Action<string> CableNameChanged;
        public event Action<string> DBConnectionNameChanged;

        public FileInfo MSWordFile
        {
            get { return new FileInfo(filePathEntry.Text); }
        }

        public string DBConnectionName
        {
            get { return connectionCombobox.ActiveText; }
        }

        public MainParseWindow() : this(new Builder("MainParseWindow.glade")) { }


        public void SetCablesNames(string[] cablesNames)
        {
            cableTypeCombobox.AppendTextCollection(cablesNames);
        }

        public void SetDBConnectionsNames(string[] connectionsNames)
        {
            connectionCombobox.AppendTextCollection(connectionsNames);
        }

        public void UpdateProgress(int parseOperationsCount, int completedOperationsCount)
        {
            //throw new NotImplementedException();
        }

        public void ParseFinishReport()
        {
            filePathEntry.Sensitive = true;
            parseButton.Sensitive = true;
            openDocButton.Sensitive = true;
            connectionCombobox.Sensitive = true;
            cableTypeCombobox.Sensitive = true;
        }

        private MainParseWindow(Builder builder) : base(builder.GetObject("MainParseWindow").Handle)
        {
            builder.Autoconnect(this);

            Initialize();

            this.DeleteEvent += Window_DeleteEvent;
            openDocButton.Clicked += OpenDocButton_Clicked;
            parseButton.Clicked += ParseButton_Clicked;
            connectionCombobox.Changed += ConnectionCombobox_Changed;
            cableTypeCombobox.Changed += CableTypeCombobox_Changed;

            var context = SynchronizationContext.Current;
        }

        private void Initialize()
        {
            chooserDialog = new FileChooserDialog("Выберите docx файл для разбора", this, FileChooserAction.Open);
            chooserDialog.AddButton("Открыть", ResponseType.Ok);
            chooserDialog.AddButton("Отмена", ResponseType.Cancel);
            var filterDocx = new FileFilter() { Name = "Файлы Microsoft Word *.docx" };
            filterDocx.AddPattern("*.docx");
            var filterAll = new FileFilter() { Name = "Все файлы" };
            filterAll.AddPattern("*");
            chooserDialog.AddFilter(filterDocx);
            chooserDialog.AddFilter(filterAll);
        }

        private void CableTypeCombobox_Changed(object sender, EventArgs e)
        {
            var cableName = cableTypeCombobox.ActiveText;
            CableNameChanged?.Invoke(cableName);
        }

        private void ConnectionCombobox_Changed(object sender, EventArgs e)
        {
            var dbConnectionName = connectionCombobox.ActiveText;
            DBConnectionNameChanged?.Invoke(dbConnectionName);
        }

        private void ParseButton_Clicked(object sender, EventArgs e)
        {
            filePathEntry.Sensitive = false;
            parseButton.Sensitive = false;
            openDocButton.Sensitive = false;
            connectionCombobox.Sensitive = false;
            cableTypeCombobox.Sensitive = false;

            TableParseStarted?.Invoke();
        }

        private void OpenDocButton_Clicked(object sender, EventArgs e)
        {
            var response = chooserDialog.Run();

            if ((ResponseType)response == ResponseType.Ok)
                filePathEntry.Text = chooserDialog.File.Path;

            chooserDialog.Hide();
        }

        private void Window_DeleteEvent(object sender, DeleteEventArgs e)
        {
            chooserDialog.Dispose();
            Application.Quit();
        }
    }
}
