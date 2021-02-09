using System;
using System.IO;

namespace GetInfoFromWordToFireBirdTable
{
    internal interface IView
    {
        string DBConnectionName { get; } 
        FileInfo MSWordFile { get; }

        void SetCablesNames(string[] cablesNames);
        void SetDBConnectionsNames(string[] connectionsNames);
        void UpdateProgress(int parseOperationsCount, int completedOperationsCount);
        void ParseFinishReport();

        event Action TableParseStarted;
        event Action<string> CableNameChanged;
        event Action<string> DBConnectionNameChanged;
    }
}
