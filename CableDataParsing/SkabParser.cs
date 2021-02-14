using System.Collections.Generic;
using WordObj = Microsoft.Office.Interop.Word;
using System.IO;
using System;
using Cables;
using FirebirdDatabaseProvider;
using System.Linq;
using CableDataParsing.MSWordTableParsers;
using CableDataParsing.TableEntityes;
using System.Text;
using CableDataParsing.NameBuilders;

namespace CableDataParsing
{
    public class SkabParser : ICableDataParcer
    {
        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;

        private string _dbConnectionString;
        private FirebirdDBProvider _dBProvider;
        private FirebirdDBTableProvider<SkabPresenter> _skabTableProvider;
        private FirebirdDBTableProvider<ListCableProperties> _ListCablePropertiesProvider;

        public event Action<int, int> ParseReport;
        public SkabParser(string dbConnectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _dbConnectionString = dbConnectionString;
            _dBProvider = new FirebirdDBProvider(_dbConnectionString);
            _skabTableProvider = new FirebirdDBTableProvider<SkabPresenter>(_dBProvider);
            _ListCablePropertiesProvider = new FirebirdDBTableProvider<ListCableProperties>(_dBProvider);
        }
        public int ParseDataToDatabase()
        {
            _dBProvider.OpenConnection();
            if (!_skabTableProvider.TableExists())
            {
                _dBProvider.CloseConnection();
                throw new Exception($"Table \"{_skabTableProvider.TableName}\",associated with {typeof(SkabPresenter)}, is not exists!");
            }

            int recordsCount = 0;

            var app = new WordObj.Application { Visible = false };
            object fileName = _mSWordFile.FullName;

            try
            {
                app.Documents.Open(ref fileName);
                var document = app.ActiveDocument;
                var tables = document.Tables;

                if (tables.Count > 0)
                {
                    var maxDiamTableCount = tables.Count / 2;
                    int tableNumber = 1;
                    _wordTableParser = new WordTableParser
                    {
                        DataRowsCount = 5,
                        RowHeadersColumnIndex = 2,
                        DataStartColumnIndex = 3,
                    };

                    var skabModifycationsList = new List<(bool HasWaterblockStripe, bool HasFilling, bool HasBraidShield)>
                    {
                        (false, true, true),
                        (true, true, true),
                        (false, false, true),
                        (true, false, true),
                        (false, true, false),
                        (true, true, false),
                        (false, false, false),
                        (true, false, false)
                    };

                    var cableShortNamesId = new List<int> { 1, 4 }; // 1 - СКАБ 250, 2 - СКАБ 660 

                    var hasArmourList = new List<(bool hasArmourBraid, bool hasArmourTube)>
                    {
                        (hasArmourBraid: false, hasArmourTube: false),
                        (hasArmourBraid: true, hasArmourTube: false),
                        (hasArmourBraid: true, hasArmourTube: true)
                    };

                    var insulationTypes = new List<int> { 0, 1 }; //1 = rubber, 0 = others

                    var twistTypesParamsList = new List<(TwistedElementType twistMode, bool hasIndividualFoilSHields, int dataStartRowIndex, int dataColumnsCount, int ColumnHeadersRowIndex)>
                    {
                        (TwistedElementType.single, false, 4, 13, 3),
                        (TwistedElementType.pair, false, 11, 14, 10),
                        (TwistedElementType.triple, false, 16, 14, 10),
                        (TwistedElementType.pair, true, 22, 13, 21),
                        (TwistedElementType.triple, true, 27, 13, 21)
                    };

                    var plasticInsMaterialParams = new List<(int fireProtectID, int insPolymerGroupId, int coverPolymerGroupId)>
                    {
                        (7, 6, 6), /*(31, 7, 7), */(12, 4, 4) //закомментировал СКАБ LTx
                    };
                    var rubberInsMaterialParams = new List<(int fireProtectID, int insPolymerGroupId, int coverPolymerGroupId)>
                    {
                        (17, 3, 6), /*(41, 8, 7), */(22, 3, 4), (24, 3, 5) //закомментировал СКАБ LTx
                    };

                    var exiParams = new List<bool> { false, true };

                    var skab = new SkabPresenter
                    {
                        TechCondId = 17,
                        HasFoilShield = true,
                        OperatingVoltageId =  //Записать
                    };

                    var billets = GetInsulatedBillets();
                    var conductors = GetConductors();

                    var listCableProperties = new ListCableProperties();
                    long recordId;

                    var stringBuilder = new StringBuilder();
                    var nameBuilder = new SkabNameBuilder(stringBuilder);

                    var skabBoolPropertyesList = new List<(bool hasProp, CableProperty propType)>();
                    while (tableNumber < maxDiamTableCount)
                    {
                        foreach(var mod in skabModifycationsList)
                        {
                            foreach (var cableShortNameId in cableShortNamesId)
                            {
                                foreach (var insType in insulationTypes)
                                {
                                    foreach (var armourType in hasArmourList)
                                    {
                                        var table = tables[tableNumber];
                                        if (table.Rows.Count > 0 && table.Columns.Count > 0)
                                        {
                                            List<TableCellData> tableData;
                                            foreach (var twistTypeParams in twistTypesParamsList)
                                            {
                                                _wordTableParser.DataColumnsCount = twistTypeParams.dataColumnsCount;
                                                _wordTableParser.ColumnHeadersRowIndex = twistTypeParams.ColumnHeadersRowIndex;
                                                _wordTableParser.DataStartRowIndex = twistTypeParams.dataStartRowIndex;

                                                tableData = _wordTableParser.GetCableCellsCollection(table);
                                                List<(int fireProtectID, int insPolymerGroupId, int coverPolymerGroupId)> materialParams;
                                                foreach (var tableCellData in tableData)
                                                {
                                                    if (int.TryParse(tableCellData.ColumnHeaderData, out int elementsCount) &&
                                                        decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                                                        decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                                                    {
                                                        materialParams = insType == 0 ? plasticInsMaterialParams : rubberInsMaterialParams;
                                                        foreach (var matParam in materialParams)
                                                        {
                                                            ConductorPresenter conductor;
                                                            foreach (var exiParam in exiParams)
                                                            {
                                                                conductor = conductors.Where(c => c.MetalId == 2 &&
                                                                                               c.Class == 2 &&
                                                                                               c.AreaInSqrMm == conductorAreaInSqrMm).First();

                                                                skab.BilletId = billets.Where(b => b.CableShortNameId == cableShortNameId &&
                                                                                                   b.PolymerGroupId == matParam.insPolymerGroupId &&
                                                                                                   b.ConductorId == conductor.ConductorId)
                                                                                       .First().BilletId;

                                                                skab.ElementsCount = elementsCount;
                                                                skab.TwistedElementTypeId = (int)twistTypeParams.twistMode;
                                                                skab.MaxCoverDiameter = maxCoverDiameter;
                                                                skab.FireProtectionId = matParam.fireProtectID;
                                                                skab.CoverPolimerGroupId = matParam.coverPolymerGroupId;
                                                                skab.HasIndividualFoilShields = twistTypeParams.hasIndividualFoilSHields;
                                                                skabBoolPropertyesList.Add((skab.HasIndividualFoilShields, CableProperty.HasIndividualFoilShields));
                                                                skabBoolPropertyesList.Add((skab.HasFoilShield, CableProperty.HasFoilShield));
                                                                skab.HasBraidShield = mod.HasBraidShield;
                                                                skabBoolPropertyesList.Add((skab.HasBraidShield, CableProperty.HasBraidShield));
                                                                skab.HasFilling = mod.HasFilling;
                                                                skabBoolPropertyesList.Add((skab.HasFilling, CableProperty.HasFilling));
                                                                skab.HasArmourBraid = armourType.hasArmourBraid;
                                                                skabBoolPropertyesList.Add((skab.HasArmourBraid, CableProperty.HasArmourBraid));
                                                                skab.HasArmourTube = armourType.hasArmourTube;
                                                                skabBoolPropertyesList.Add((skab.HasArmourTube, CableProperty.HasArmourTube));
                                                                skab.HasWaterBlockStripe = mod.HasWaterblockStripe;
                                                                skabBoolPropertyesList.Add((skab.HasWaterBlockStripe, CableProperty.HasWaterBlockStripe));
                                                                skab.SparkSafety = exiParam;
                                                                skabBoolPropertyesList.Add((skab.SparkSafety, CableProperty.SparkSafety));
                                                                skab.CoverColorId = (exiParam && armourType.hasArmourTube == false) ? 3 : 2;
                                                                skab.ClimaticModId = ;  //Записать


                                                                skab.Title = nameBuilder.GetCableName(skab, conductor: conductor);

                                                                recordId = _skabTableProvider.AddItem(skab);

                                                                listCableProperties.CableId = recordId;

                                                                foreach(var boolPropPair in skabBoolPropertyesList)
                                                                {
                                                                    if (boolPropPair.hasProp)
                                                                    {
                                                                        listCableProperties.PropertyId = (long)boolPropPair.propType;
                                                                        _ListCablePropertiesProvider.AddItem(listCableProperties);
                                                                    }
                                                                }

                                                                skabBoolPropertyesList.Clear();
                                                                recordsCount++;
                                                            }
                                                        }
                                                    }
                                                    else throw new Exception($"Не удалось распарсить ячейку таблицы №{tableNumber}!");
                                                }
                                                tableData.Clear();
                                            }
                                        }
                                        else throw new Exception("Таблица пуста!");
                                        ParseReport?.Invoke(maxDiamTableCount, tableNumber);
                                        tableNumber++;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                _dBProvider.CloseConnection();
                app.Quit();
            }
            return recordsCount;
        }

        private ICollection<CableBilletPresenter> GetInsulatedBillets()
        {
            var billetProvider = new FirebirdDBTableProvider<CableBilletPresenter>(_dBProvider);
            var result = billetProvider.GetAllItemsFromTable();
            return result;
        }

        private ICollection<ConductorPresenter> GetConductors()
        {
            var conductorProvider = new FirebirdDBTableProvider<ConductorPresenter>(_dBProvider);
            var result = conductorProvider.GetAllItemsFromTable();
            return result;
        }

    }
}
