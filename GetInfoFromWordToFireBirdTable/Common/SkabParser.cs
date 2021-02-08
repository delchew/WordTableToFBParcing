using System.Collections.Generic;
using WordObj = Microsoft.Office.Interop.Word;
using System.IO;
using System;
using Cables;
using System.Text;
using GetInfoFromWordToFireBirdTable.Database.Extensions;
using GetInfoFromWordToFireBirdTable.CableEntityes;
using System.Linq;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class SkabParser : ICableDataParcer
    {
        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private FirebirdDBTableProvider<CableEntityes.Skab> _skabTableProvider;
        private StringBuilder _stringBuilder = new StringBuilder();

        private Random random = new Random(); //TODO Удалить нахер потом эту шляпу!

        public event Action<int, int> ParseReport;
        public SkabParser(FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
        }
        public int ParseDataToDatabase()
        {
            _skabTableProvider = new FirebirdDBTableProvider<CableEntityes.Skab>();
            _skabTableProvider.OpenConnection();
            if (!_skabTableProvider.TableExists())
            {
                _skabTableProvider.CloseConnection();
                throw new Exception($"Table \"{_skabTableProvider.TableName}\",associated with {typeof(CableEntityes.Skab)}, is not exists!");
            }

            int recordsCount = 0;

            //var app = new WordObj.Application { Visible = false };
            //object fileName = _mSWordFile.FullName;

            try
            {
                //app.Documents.Open(ref fileName);
                //var document = app.ActiveDocument;
                //var tables = document.Tables;

                if (true/*tables.Count > 0*/)
                {
                    var maxDiamTableCount = 96;//tables.Count / 2;
                    int tableNumber = 1;
                    //_wordTableParser = new WordTableParser
                    //{
                    //    DataRowsCount = 5,
                    //    RowHeadersColumnIndex = 2,
                    //    DataStartColumnIndex = 3,
                    //};

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

                    var skabVoltageTypes = new List<int> { 250, 660 };

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

                    var skab = new CableEntityes.Skab { TechCondId = 17, HasFoilShield = true };

                    var billets = GetInsulatedBillets();

                    while (tableNumber < maxDiamTableCount)
                    {
                        foreach(var mod in skabModifycationsList)
                        {
                            foreach (var voltageType in skabVoltageTypes)
                            {
                                foreach (var insType in insulationTypes)
                                {
                                    foreach (var armourType in hasArmourList)
                                    {
                                        ///var table = tables[tableNumber];
                                        if (true/*table.Rows.Count > 0 && table.Columns.Count > 0*/)
                                        {
                                            //List<TableCellData> tableData;
                                            foreach (var twistTypeParams in twistTypesParamsList)
                                            {
                                                //_wordTableParser.DataColumnsCount = twistTypeParams.dataColumnsCount;
                                                //_wordTableParser.ColumnHeadersRowIndex = twistTypeParams.ColumnHeadersRowIndex;
                                                //_wordTableParser.DataStartRowIndex = twistTypeParams.dataStartRowIndex;

                                                //tableData = _wordTableParser.GetCableCellsCollection(table);
                                                List<(int fireProtectID, int insPolymerGroupId, int coverPolymerGroupId)> materialParams;
                                                for (int i = 0; i < 65; i++)//foreach (var tableCellData in tableData)
                                                {
                                                    if (true/*int.TryParse(tableCellData.ColumnHeaderData, out int elementsCount) &&
                                                        double.TryParse(tableCellData.CellData, out double maxCoverDiameter) &&
                                                        double.TryParse(tableCellData.RowHeaderData, out double conductorAreaInSqrMm)*/)
                                                    {
                                                        materialParams = insType == 0 ? plasticInsMaterialParams : rubberInsMaterialParams;
                                                        foreach (var matParam in materialParams)
                                                        {
                                                            foreach (var exiParam in exiParams)
                                                            {
                                                                skab.BilletId = billets.Where(b => b.CableShortNameId == (voltageType == 250 ? 1 : 4) &&
                                                                                                   b.PolymerGroupId == matParam.insPolymerGroupId)
                                                                                       .First().BilletId;//GetInsBilletId(matParam.insPolymerGroupId, voltageType, i/*conductorAreaInSqrMm*/);
                                                                skab.ElementsCount = i;//elementsCount;
                                                                skab.TwistedElementTypeId = (int)twistTypeParams.twistMode;
                                                                skab.MaxCoverDiameter = i;//maxCoverDiameter;
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
                                                //tableData.Clear();
                                            }
                                        }
                                        else throw new Exception("Таблица пуста!");
                                        tableNumber++;
                                    }
                                }
                            }
                        }
                        ParseReport?.Invoke(maxDiamTableCount, tableNumber);
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
                //app.Quit();
            }
            return recordsCount;
        }

        private int GetInsBilletId(int insPolymerGroupId, int skabVoltageType, double conductorAreaInSqrMm)
        {
            _stringBuilder.Append(@"SELECT INS_BILLET_ID FROM INSULATED_BILLET WHERE ");
            _stringBuilder.Append($@"POLYMER_GROUP_ID = {insPolymerGroupId} AND ");
            int shortNameId = skabVoltageType == 250 ? 1 : 4;
            _stringBuilder.Append($@"SHRT_N_ID = {shortNameId} AND ");
            _stringBuilder.Append(@"COND_ID = (SELECT COND_ID FROM CONDUCTOR WHERE METAL_ID = 2 AND WIRES_COUNT > 1 AND ");
            _stringBuilder.Append($@"AREA_IN_SQR_MM = {conductorAreaInSqrMm.ToFBSqlString()});");
            var sqlRequest = _stringBuilder.ToString();
            _stringBuilder.Clear();
            var result = _skabTableProvider.GetSingleObjBySQL(sqlRequest);
            if (result != null)
            {
                var billetId = (int)result;
                return billetId;
            }
            throw new NullReferenceException("SQL запрос вернул null!");
        }

        private ICollection<CableBillet> GetInsulatedBillets()
        {
            var billetProvider = new FirebirdDBTableProvider<CableBillet>();
            billetProvider.OpenConnection();
            var result = billetProvider.GetAllItemsFromTable();
            billetProvider.CloseConnection();
            return result;

        }
    }
}
