using System;
using System.IO;

namespace GuiPresenter
{
    public interface IView
    {
        string DBConnectionName { get; } 
        FileInfo MSWordFile { get; }

        void SetCablesNames(string[] cablesNames);
        void SetDBConnectionsNames(string[] connectionsNames);
        void UpdateProgress(double completedPersentage);
        void ParseFinishReport();

        event Action TableParseStarted;
        event Action<string> CableNameChanged;
        event Action<string> DBConnectionNameChanged;
    }
}
