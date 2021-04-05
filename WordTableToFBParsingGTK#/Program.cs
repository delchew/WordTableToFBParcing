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

            var win = new MainParseWindow();
            //var messageService = new MessageService();
            var presenter = new MainPresenter(win, null);
            app.AddWindow(win);

            win.Show();
            Application.Run();
        }
    }
}
