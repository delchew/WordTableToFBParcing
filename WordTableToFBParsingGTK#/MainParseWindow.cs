using System;
using System.IO;
using Gtk;
using GuiPresenter;
using UI = Gtk.Builder.ObjectAttribute;

namespace WordTableToFBParsingGTK
{
    class MainParseWindow : Window, IView
    {
        [UI] private Entry filePathEntry = null;
        [UI] private Button openDocButton = null;
        [UI] private ComboBoxText connectionCombobox = null;
        [UI] private ComboBoxText cableTypeCombobox = null;
        [UI] private Button parseButton = null;
        [UI] private ProgressBar progressBar = null;

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
            throw new NotImplementedException();
        }

        public void ParseFinishReport()
        {
            throw new NotImplementedException();
        }

        private MainParseWindow(Builder builder) : base(builder.GetObject("MainParseWindow").Handle)
        {
            builder.Autoconnect(this);

            DeleteEvent += Window_DeleteEvent;

            openDocButton.Clicked += OpenDocButton_Clicked;
            parseButton.Clicked += ParseButton_Clicked;
            connectionCombobox.Changed += ConnectionCombobox_Changed;
            cableTypeCombobox.Changed += CableTypeCombobox_Changed;

            progressBar.PulseStep = 0.1;
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
            TableParseStarted?.Invoke();
        }

        private void OpenDocButton_Clicked(object sender, EventArgs e)
        {
            var openButton = new FileChooserButton("Открыть", FileChooserAction.Open);
            var chooserDialog = new FileChooserDialog("Выберите docx файл для разбора", this, FileChooserAction.Open, openButton);
            //chooserDialog.Show();
            chooserDialog.Run(); //TODO

            filePathEntry.Text = chooserDialog.File.Path;
        }

        private void Window_DeleteEvent(object sender, DeleteEventArgs a)
        {
            Application.Quit();
        }
    }
}
