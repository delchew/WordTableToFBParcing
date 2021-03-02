using System;
using System.Collections.Generic;
using System.IO;
using WordObj = Microsoft.Office.Interop.Word;
using FirebirdDatabaseProvider;
using CableDataParsing.MSWordTableParsers;
using Cables.Common;
using CableDataParsing.TableEntityes;
using System.Linq;
using CableDataParsing.NameBuilders;
using System.Text;
using Cables;

namespace CableDataParsing
{
    public class KunrsParser : ICableDataParcer
    {
        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _dbConnectionString;
        private FirebirdDBTableProvider<KunrsPresenter> _kunrsTableProvider;
        private FirebirdDBTableProvider<ListCablePowerColor> _ListCablePowerColorProvider;
        private FirebirdDBTableProvider<ListCableProperties> _ListCablePropertiesProvider;
        private FirebirdDBTableProvider<ListCableBillets> _ListCableBilletsProvider;
        private FirebirdDBProvider _dBProvider;

        public KunrsParser(string dbConnectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _dbConnectionString = dbConnectionString;
            _dBProvider = new FirebirdDBProvider(_dbConnectionString);
            _kunrsTableProvider = new FirebirdDBTableProvider<KunrsPresenter>(_dBProvider);
            _ListCablePowerColorProvider = new FirebirdDBTableProvider<ListCablePowerColor>(_dBProvider);
            _ListCablePropertiesProvider = new FirebirdDBTableProvider<ListCableProperties>(_dBProvider);
            _ListCableBilletsProvider = new FirebirdDBTableProvider<ListCableBillets>(_dBProvider);
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
                        { 0, 6 }, { 1, 4 }, { 2, 5 }
                    };
                    var insBilletKunrsIdDictionary = new Dictionary<decimal, int>
                    {
                        { 0.75m, 1 }, { 1.0m, 2 }, { 1.5m, 3 }, { 2.5m, 4 }, { 4.0m, 5 }, { 6.0m, 6 }, { 10.0m, 7 }, { 16.0m, 8 }
                    };
                    var powerColorsDict = new Dictionary<decimal, PowerWiresColorScheme[]>
                    {
                        { 2m, new [] { PowerWiresColorScheme.N } },
                        { 3m, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none} },
                        { 4m, new [] { PowerWiresColorScheme.N, PowerWiresColorScheme.PE } },
                        { 5m, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none } }
                    };

                    var dictCablePropsToId = new Dictionary<CableProperty, long>
                    {
                        { CableProperty.HasIndividualFoilShields, 1L },
                        { CableProperty.HasIndividualBraidShield, 2L },
                        { CableProperty.HasFoilShield, 3L },
                        { CableProperty.HasBraidShield, 4L },
                        { CableProperty.HasFilling, 5L },
                        { CableProperty.HasArmourBraid, 6L },
                        { CableProperty.HasArmourTape, 7L },
                        { CableProperty.HasArmourTube, 8L },
                        { CableProperty.HasWaterBlockStripe, 9L },
                        { CableProperty.SparkSafety, 10L }
                    };

                    var kunrs = new KunrsPresenter
                    {
                        TwistedElementTypeId = 1,
                        TechCondId = 25,
                        HasFilling = true,
                        OperatingVoltageId = 1,
                        ClimaticModId = 3
                    };

                    var conductors = GetConductors();

                    var listCablePowerColor = new ListCablePowerColor();
                    var listCableProperties = new ListCableProperties();
                    var listCableBillets = new ListCableBillets();
                    long recordId;
                    ConductorPresenter conductor;
                    PowerWiresColorScheme[] powerColorSchemeArray;

                    var stringBuilder = new StringBuilder();
                    var nameBuilder = new KunrsNameBuider(stringBuilder);

                    var kunrsBoolPropertyesList = new List<(bool hasProp, CableProperty propType)>();

                    for (int i = 0; i < hasFoilShieldDictionary.Count; i++)
                    {
                        var tableData = _wordTableParser.GetCableCellsCollection(table);
                        foreach (var tableCellData in tableData)
                        {
                            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                                decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                            {
                                for (int j = 0; j < polimerGroupIdDictionary.Count; j++)
                                {
                                    powerColorSchemeArray = powerColorsDict[elementsCount];
                                    for (int k = 0; k < powerColorSchemeArray.Length; k++)
                                    {
                                        kunrs.ElementsCount = elementsCount;
                                        kunrs.MaxCoverDiameter = maxCoverDiameter;
                                        kunrs.FireProtectionId = polimerGroupIdDictionary[j] == 6 ? 18 : 23;
                                        kunrs.CoverPolimerGroupId = polimerGroupIdDictionary[j];
                                        kunrsBoolPropertyesList.Add((kunrs.HasFilling, CableProperty.HasFilling));
                                        kunrs.HasFoilShield = hasFoilShieldDictionary[i];
                                        kunrsBoolPropertyesList.Add((kunrs.HasFoilShield, CableProperty.HasFoilShield));
                                        kunrs.HasArmourBraid = hasArmourDictionary[i];
                                        kunrsBoolPropertyesList.Add((kunrs.HasArmourBraid, CableProperty.HasArmourBraid));
                                        kunrs.HasArmourTube = hasArmourDictionary[i];
                                        kunrsBoolPropertyesList.Add((kunrs.HasArmourTube, CableProperty.HasArmourTube));
                                        kunrs.PowerColorSchemeId = (int)powerColorSchemeArray[k];
                                        kunrs.CoverColorId = polimerGroupIdDictionary[j] == 5 ? 8 : 2;

                                        conductor = conductors.Where(c => c.MetalId == 3 &&
                                                                          (c.Class == 2 || c.Class == 3) &&
                                                                          c.AreaInSqrMm == conductorAreaInSqrMm).First();
                                        kunrs.Title = nameBuilder.GetCableName(kunrs, conductor: conductor, parameter: powerColorSchemeArray[k].GetDescription());

                                        recordId = _kunrsTableProvider.AddItem(kunrs);

                                        listCablePowerColor.CableId = recordId;
                                        listCablePowerColor.PowerColorSchemeId = kunrs.PowerColorSchemeId;
                                        _ListCablePowerColorProvider.AddItem(listCablePowerColor);

                                        listCableBillets.CableId = recordId;
                                        listCableBillets.BilletId = insBilletKunrsIdDictionary[conductorAreaInSqrMm];
                                        _ListCableBilletsProvider.AddItem(listCableBillets);

                                        listCableProperties.CableId = recordId;

                                        foreach (var boolPropPair in kunrsBoolPropertyesList)
                                        {
                                            if (boolPropPair.hasProp)
                                            {
                                                listCableProperties.PropertyId = dictCablePropsToId[boolPropPair.propType];
                                                _ListCablePropertiesProvider.AddItem(listCableProperties);
                                            }
                                        }

                                        kunrsBoolPropertyesList.Clear();
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

        private ICollection<ConductorPresenter> GetConductors()
        {
            var conductorProvider = new FirebirdDBTableProvider<ConductorPresenter>(_dBProvider);
            var result = conductorProvider.GetAllItemsFromTable();
            return result;
        }
    }
}
