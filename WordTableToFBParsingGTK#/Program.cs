using System;
using Gtk;
using GuiPresenter;

namespace WordTableToFBParsingGTK
{
    class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            Application.Init();

            var app = new Application("org.WordTableToFBParsingGTK_.WordTableToFBParsingGTK_", GLib.ApplicationFlags.None);
            app.Register(GLib.Cancellable.Current);

            var window = new MainParseWindow();
            var messageDialog = new MessageService(window);
            new MainPresenter(window, messageDialog);
            app.AddWindow(window);

            window.Show();
            Application.Run();
        }
    }
}
