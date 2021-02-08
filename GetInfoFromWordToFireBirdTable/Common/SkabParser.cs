﻿using System.Collections.Generic;
using WordObj = Microsoft.Office.Interop.Word;
using System.IO;
using System;
using Cables;
using System.Text;
using FirebirdDatabaseProvider;
using GetInfoFromWordToFireBirdTable.CableEntityes;
using System.Linq;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class SkabParser : ICableDataParcer
    {
        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _dbConnectionString;
        private FirebirdDBTableProvider<Skab> _skabTableProvider;
        private StringBuilder _stringBuilder = new StringBuilder();

        public event Action<int, int> ParseReport;
        public SkabParser(string dbConnectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _dbConnectionString = dbConnectionString;
        }
        public int ParseDataToDatabase()
        {
            _skabTableProvider = new FirebirdDBTableProvider<Skab>(_dbConnectionString);
            _skabTableProvider.OpenConnection();
            if (!_skabTableProvider.TableExists())
            {
                _skabTableProvider.CloseConnection();
                throw new Exception($"Table \"{_skabTableProvider.TableName}\",associated with {typeof(Skab)}, is not exists!");
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

                    var skab = new Skab { TechCondId = 17, HasFoilShield = true };

                    var billets = GetInsulatedBillets();
                    var conductors = GetConductors();

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
                                                            Conductor conductor;
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
                                                                skab.SparkSafety = exiParam;
                                                                skab.CoverColorId = (exiParam && armourType.hasArmourTube == false) ? 3 : 2;
                                                                skab.HasIndividualFoilShields = twistTypeParams.hasIndividualFoilSHields;
                                                                skab.HasBraidShield = mod.HasBraidShield;
                                                                skab.HasFilling = mod.HasFilling;
                                                                skab.HasArmourBraid = armourType.hasArmourBraid;
                                                                skab.HasArmourTube = armourType.hasArmourTube;
                                                                skab.HasWaterBlockStripe = mod.HasWaterblockStripe;

                                                                _skabTableProvider.AddItem(skab);
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
                _skabTableProvider.CloseConnection();
                app.Quit();
            }
            return recordsCount;
        }

        private ICollection<CableBillet> GetInsulatedBillets()
        {
            var billetProvider = new FirebirdDBTableProvider<CableBillet>(_dbConnectionString);
            billetProvider.OpenConnection();
            var result = billetProvider.GetAllItemsFromTable();
            billetProvider.CloseConnection();
            return result;
        }

        private ICollection<Conductor> GetConductors()
        {
            var conductorProvider = new FirebirdDBTableProvider<Conductor>(_dbConnectionString);
            conductorProvider.OpenConnection();
            var result = conductorProvider.GetAllItemsFromTable();
            conductorProvider.CloseConnection();
            return result;
        }

    }
}
