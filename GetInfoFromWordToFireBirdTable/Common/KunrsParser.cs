using GetInfoFromWordToFireBirdTable.TableEntityes;
using System;
using System.Collections.Generic;
using System.IO;
using WordObj = Microsoft.Office.Interop.Word;
using FirebirdDatabaseProvider;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class KunrsParser : ICableDataParcer
    {
        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _dbConnectionString;
        private FirebirdDBTableProvider<KunrsPresenter> _kunrsTableProvider;
        private FirebirdDBTableProvider<CableIdUnionPowerColorSchemeIdPresenter> _kunrsPowerSchemeUnionProvider;
        private FirebirdDBProvider _dBProvider;

        public KunrsParser(string dbConnectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _dbConnectionString = dbConnectionString;
            _dBProvider = new FirebirdDBProvider(_dbConnectionString);
            _kunrsTableProvider = new FirebirdDBTableProvider<KunrsPresenter>(_dBProvider);
            _kunrsPowerSchemeUnionProvider = new FirebirdDBTableProvider<CableIdUnionPowerColorSchemeIdPresenter>(_dBProvider);
        }

        public event Action<int, int> ParseReport;
        public int ParseDataToDatabase()
        {

            _dBProvider.OpenConnection();
            if(!_kunrsTableProvider.TableExists())
            {
                _dBProvider.CloseConnection();
                throw new Exception($"Table \"{_kunrsTableProvider.TableName}\",associated with {typeof(KunrsPresenter)}, is not exists!");
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
                    var insBilletKunrsIdDictionary = new Dictionary<decimal, int>
                    {
                        { 0.75m, 8 }, { 1.0m, 7 }, { 1.5m, 6 }, { 2.5m, 5 }, { 4.0m, 4 }, { 6.0m, 3 }, { 10.0m, 2 }, { 16.0m, 1 }
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
                                decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                            {
                                for (int j = 0; j < polimerGroupIdDictionary.Count; j++)
                                {
                                    var powerColorSchemeArray = powerColorsDict[elementsCount];
                                    for (int k = 0; k < powerColorSchemeArray.Length; k++)
                                    {
                                        var kunrs = new KunrsPresenter
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

                                        var cableUnionPowerColorScheme = new CableIdUnionPowerColorSchemeIdPresenter
                                        {
                                            CableId = 1,
                                            PowerColorSchemeId = (int)powerColorSchemeArray[k]
                                        };

                                        _kunrsTableProvider.AddItem(kunrs);
                                        _kunrsPowerSchemeUnionProvider.AddItem(cableUnionPowerColorScheme);
                                        recordsCount++;
                                    }
                                }
                            }
                        }
                        ParseReport?.Invoke(hasFoilShieldDictionary.Count, i + 1);
                        _wordTableParser.DataStartRowIndex += _wordTableParser.DataRowsCount;
                        tableData.Clear();
                    }
                }
            }
            finally
            {
                app.Quit();
                _dBProvider.CloseConnection();
            }
            return recordsCount;
        }
    }
}
