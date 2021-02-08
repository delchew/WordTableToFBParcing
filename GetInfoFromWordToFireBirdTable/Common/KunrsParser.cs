using GetInfoFromWordToFireBirdTable.CableEntityes;
using System;
using System.Collections.Generic;
using System.IO;
using WordObj = Microsoft.Office.Interop.Word;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class KunrsParser : ICableDataParcer
    {
        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private FirebirdDBTableProvider<Kunrs> _kunrsTableProvider;
        public KunrsParser(FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _kunrsTableProvider = new FirebirdDBTableProvider<Kunrs>();
        }

        public event Action<int, int> ParseReport;
        public int ParseDataToDatabase()
        {
            _kunrsTableProvider.OpenConnection();
            if(!_kunrsTableProvider.TableExists())
            {
                _kunrsTableProvider.CloseConnection();
                throw new Exception($"Table \"{_kunrsTableProvider.TableName}\",associated with {typeof(Kunrs)}, is not exists!");
            }

            int recordsCount = 0;

            var app = new WordObj.Application { Visible = false };
            object fileName = _mSWordFile.FullName;

            try
            {
                app.Documents.Open(ref fileName);
                var document = app.ActiveDocument;
                var table = document.Tables[1];

                if (table.Rows.Count > 0 && table.Columns.Count > 0)
                {
                    _wordTableParser = new WordTableParser
                    {
                        DataColumnsCount = 4,
                        DataRowsCount = 8,
                        ColumnHeadersRowIndex = 3,
                        RowHeadersColumnIndex = 2,
                        DataStartColumnIndex = 3,
                        DataStartRowIndex = 4
                    };

                    var hasFoilShieldDictionary = new Dictionary<int, bool>
                    {
                        { 0, false }, { 1, true }, { 2, false }, { 3, true }
                    };
                    var hasArmourDictionary = new Dictionary<int, bool>
                    {
                        { 0, false }, { 1, false }, { 2, true }, { 3, true }
                    };
                    var polimerGroupIdDictionary = new Dictionary<int, int>
                    {
                        { 0, 1 }, { 1, 4 }, { 2, 5 }
                    };
                    var insBilletKunrsIdDictionary = new Dictionary<double, int>
                    {
                        { 0.75, 8 }, { 1.0, 7 }, { 1.5, 6 }, { 2.5, 5 }, { 4.0, 4 }, { 6.0, 3 }, { 10.0, 2 }, { 16.0, 1 }
                    };
                    var powerColorsDict = new Dictionary<int, PowerWiresColorScheme[]>
                    {
                        { 2, new [] { PowerWiresColorScheme.N } },
                        { 3, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none} },
                        { 4, new [] { PowerWiresColorScheme.N, PowerWiresColorScheme.PE } },
                        { 5, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none } }
                    };

                    for (int i = 0; i < hasFoilShieldDictionary.Count; i++)
                    {
                        var tableData = _wordTableParser.GetCableCellsCollection(table);
                        foreach (var tableCellData in tableData)
                        {
                            if (int.TryParse(tableCellData.ColumnHeaderData, out int elementsCount) &&
                                double.TryParse(tableCellData.CellData, out double maxCoverDiameter) &&
                                double.TryParse(tableCellData.RowHeaderData, out double conductorAreaInSqrMm))
                            {
                                for (int j = 0; j < polimerGroupIdDictionary.Count; j++)
                                {
                                    var powerColorSchemeArray = powerColorsDict[elementsCount];
                                    for (int k = 0; k < powerColorSchemeArray.Length; k++)
                                    {
                                        var kunrs = new Kunrs
                                        {
                                            BilletId = insBilletKunrsIdDictionary[conductorAreaInSqrMm],
                                            ElementsCount = elementsCount,
                                            TwistedElementTypeId = 1,
                                            TechCondId = 26,
                                            MaxCoverDiameter = maxCoverDiameter,
                                            FireProtectionId = polimerGroupIdDictionary[j] == 1 ? 17 : 22,
                                            CoverPolimerGroupId = polimerGroupIdDictionary[j],
                                            HasFoilShield = hasFoilShieldDictionary[i],
                                            HasFilling = true,
                                            HasArmourBraid = hasArmourDictionary[i],
                                            HasArmourTube = hasArmourDictionary[i],
                                            PowerColorSchemeId = (int)powerColorSchemeArray[k],
                                            CoverColorId = polimerGroupIdDictionary[j] == 5 ? 8 : 2
                                        };
                                        _kunrsTableProvider.AddItem(kunrs);
                                        recordsCount++;
                                    }
                                }
                            }
                        }
                        ParseReport?.Invoke(hasFoilShieldDictionary.Count, i + 1);
                        _wordTableParser.DataStartRowIndex += _wordTableParser.DataRowsCount;
                        tableData.Clear();
                    }
                    _kunrsTableProvider.CloseConnection();
                }
            }
            finally
            {
                app.Quit();
            }
            return recordsCount;
        }
    }
}
