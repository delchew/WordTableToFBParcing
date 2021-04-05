using GuiPresenter;
using System;
using System.Windows.Forms;

namespace WordTableToFBParsing
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            var mainForm = new MainForm();
            var messageService = new MessageService();
            var presenter = new MainPresenter(mainForm, messageService);

            Application.Run(mainForm);
        }
    }
}
