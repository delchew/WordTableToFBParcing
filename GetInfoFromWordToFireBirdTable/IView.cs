using System;
using System.IO;

namespace GetInfoFromWordToFireBirdTable
{
    public interface IView
    {
        FileInfo FBDatabaseFile { get; } 
        FileInfo MSWordFile { get; }

        void SetCablesNames(string[] cablesNames);
        void UpdateProgress(int parseOperationsCount, int completedOperationsCount);
        void ParseFinishReport();

        event Action TableParseStarted;
        event Action<string> CableNameChanged;
    }
}
