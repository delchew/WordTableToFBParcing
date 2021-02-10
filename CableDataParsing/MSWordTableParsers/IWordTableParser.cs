using System.Collections.Generic;
using WordObj = Microsoft.Office.Interop.Word;

namespace CableDataParsing.MSWordTableParsers
{
    public interface IWordTableParser<T> where T : new()
    {
        List<T> GetCableCellsCollection(WordObj.Table table);
    }
}
