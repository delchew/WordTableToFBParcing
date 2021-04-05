using GuiPresenter;
using Gtk;

namespace WordTableToFBParsingGTK
{
    public class MessageService : IMessageService
    {
        private MessageDialog _dialog;
        private Window _parentWindow;
        public MessageService(Window parentWindow)
        {
            _parentWindow = parentWindow;
        }

        public void ShowError(string error)
        {
            ModalDialogWindowWork(MessageType.Error, error);
        }

        public void ShowExclamation(string exclamation)
        {
            ModalDialogWindowWork(MessageType.Warning, exclamation);
        }

        public void ShowMessage(string message)
        {
            ModalDialogWindowWork(MessageType.Info, message);
        }

        private void ModalDialogWindowWork(MessageType messageType, string message)
        {
            using (_dialog = new MessageDialog(_parentWindow, DialogFlags.Modal, messageType, ButtonsType.Ok, message))
            {
                var response = _dialog.Run();
                if ((ResponseType)response == ResponseType.Ok)
                    _dialog.Hide();
            }
        }
    }
}
