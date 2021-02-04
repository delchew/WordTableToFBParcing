using System.Collections.Generic;
using WordObj = Microsoft.Office.Interop.Word;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public interface IWordTableParser<T> where T : new()
    {
        List<T> GetCableCellsCollection(WordObj.Table table);
    }
}
