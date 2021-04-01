using System.Collections.Generic;

namespace CableDataParsing.MSWordTableParsers
{
    public interface IWordTableParser<T> where T : new()
    {
        //List<T> GetCableCellsCollection(WordObj.Table table);
        List<T> GetCableCellsCollection(int tableIndex); 
    }
}
