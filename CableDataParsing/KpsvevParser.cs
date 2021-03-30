using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CableDataParsing.CableTitleBulders;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;
using WordObj = Microsoft.Office.Interop.Word;

namespace CableDataParsing
{
    public class KpsvevParser : CableParser
    {
        public KpsvevParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile, new KpsvevTitleBuilder())
        { }

        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;

            var app = new WordObj.Application { Visible = false };
            object fileName = _mSWordFile.FullName;

            try
            {
                app.Documents.Open(ref fileName);
                var document = app.ActiveDocument;
                var tables = document.Tables;
                if(tables.Count > 0)
                {
                    var PVCGroup = _dbContext.PolymerGroups.Find(1); // PVC
                    var PVCColdGroup = _dbContext.PolymerGroups.Find(10); // PVC Cold
                    var PVCTermGroup = _dbContext.PolymerGroups.Find(11); //PVC term
                    var PVCLSGroup = _dbContext.PolymerGroups.Find(6); //PVC LS

                    var polymerGroups = new List<PolymerGroup>
                    {
                        PVCGroup, PVCColdGroup, PVCTermGroup, PVCLSGroup
                    };

                    var PVCLSLTxGroup = _dbContext.PolymerGroups.Find(7); //PVC LSLTx
                    var PESelfExtinguish = _dbContext.PolymerGroups.Find(12); //PE self extinguish

                    var noFireProtectClass = _dbContext.FireProtectionClasses.Find(1); //О1.8.2.5.4
                    var PolymerGroupFireClassDict = new Dictionary<PolymerGroup, FireProtectionClass>()
                    {
                        { PVCGroup, noFireProtectClass },
                        { PVCColdGroup, noFireProtectClass },
                        { PVCTermGroup, noFireProtectClass },
                        { PVCLSGroup, _dbContext.FireProtectionClasses.Find(8) }, //П1б.8.2.2.2
                        { PVCLSLTxGroup, _dbContext.FireProtectionClasses.Find(28) }, //П1б.8.2.1.2
                        { PESelfExtinguish, noFireProtectClass }
                    };

                    var climaticModUHL = _dbContext.ClimaticMods.Find(3);
                    var redColor = _dbContext.Colors.Find(1);
                    var blackColor = _dbContext.Colors.Find(2);
                    var cableShortName = _dbContext.CableShortNames.Find(6); // КПСВ(Э)
                    var operatingVoltage = _dbContext.OperatingVoltages.Find(6); // 300В 50Гц, постоянка - до 500В
                    var operatingVoltageLoutoks = _dbContext.OperatingVoltages.Find(5); // 300В 50Гц
                    var TU2 = _dbContext.TechnicalConditions.Find(1);
                    var TU30 = _dbContext.TechnicalConditions.Find(35);
                    var TU49 = _dbContext.TechnicalConditions.Find(20);
                    var twistedElementType = _dbContext.TwistedElementTypes.Find(2); //pair

                    var kpsvevCableProps = new List<Cables.Common.CableProperty?>
                    {
                        null, Cables.Common.CableProperty.HasFoilShield
                    };

                    _wordTableParser = new WordTableParser
                    {
                        ColumnHeadersRowIndex = 3,
                        RowHeadersColumnIndex = 2,
                        DataRowsCount = 5,
                        DataStartColumnIndex = 3,
                    };
                    List<TableCellData> tableData1, tableData2;

                    
                    decimal? maxCoverDiameter;
                    foreach (var prop in kpsvevCableProps)
                    {
                        _wordTableParser.DataStartRowIndex = prop.HasValue ? 9 : 4;
                        _wordTableParser.DataColumnsCount = 10;
                        tableData1 = _wordTableParser.GetCableCellsCollection(tables[1]);
                        _wordTableParser.DataColumnsCount = 8;
                        tableData2 = _wordTableParser.GetCableCellsCollection(tables[2]);
                        var tableDataNoShield = tableData1.Concat(tableData2);

                        foreach (var tableCellData in tableDataNoShield)
                        {
                            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                            {

                                if (decimal.TryParse(tableCellData.CellData, out decimal diameterValue))
                                    maxCoverDiameter = diameterValue;
                                else
                                {
                                    var cableSizes = tableCellData.CellData.Split('\u00D7'); //знак умножения в юникоде
                                    if (cableSizes.Length < 2) continue;
                                    if (cableSizes.Length > 2 &&
                                        decimal.TryParse(cableSizes[0], out decimal heigth) &&
                                        decimal.TryParse(cableSizes[1], out decimal width))
                                    {
                                        maxCoverDiameter = null;

                                    }
                                    else throw new Exception("Wrong format table cell data!");

                                }

                                foreach (var polymerGroup in polymerGroups)
                                {
                                    var kpsvv = new Cable
                                    {
                                        ClimaticMod = climaticModUHL,
                                        CoverPolymerGroup = polymerGroup,
                                        TechnicalConditions = TU2,
                                        CoverColor = polymerGroup == PVCColdGroup ? blackColor : redColor,
                                        OperatingVoltage = operatingVoltage,
                                        ElementsCount = elementsCount,
                                        MaxCoverDiameter = maxCoverDiameter,
                                        FireProtectionClass = PolymerGroupFireClassDict[polymerGroup],
                                        TwistedElementType = twistedElementType,
                                    };
                                    kpsvv.Title = cableTitleBuilder.GetCableTitle(kpsvv, prop);
                                    var cableRec = _dbContext.Cables.Add(kpsvv).Entity;
                                    _dbContext.SaveChanges();

                                    if (prop.HasValue)
                                        AddCablePropertiesToDBContext(cableRec, prop.Value);
                                }
                            }
                            else
                            {
                                throw new InvalidCastException("Can't cast elementsCount or conductorAreaInSqrMm!");
                            }
                        }
                    }
                    
                    

                    

                    throw new NotImplementedException();

                }
                else
                    throw new Exception("Отсутствуют таблицы для парсинга в указанном Word файле!");
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                app.Quit();
            }

            return recordsCount;
        }

        public int ParseBillets()
        {
            var conuctor05 = _dbContext.Conductors.Find(17);
            var conuctor075 = _dbContext.Conductors.Find(18);
            var conuctor1 = _dbContext.Conductors.Find(19);
            var conuctor15 = _dbContext.Conductors.Find(20);
            var conuctor25 = _dbContext.Conductors.Find(21);

            var cableShortName = _dbContext.CableShortNames.Find(6); // КПСВ(Э)
            var operatingVoltage = _dbContext.OperatingVoltages.Find(6); // 300В 50Гц, постоянка - до 500В

            var polymerGroups = new List<PolymerGroup>
            {
                _dbContext.PolymerGroups.Find(1), // PVC
                _dbContext.PolymerGroups.Find(6), //PVC LS
                _dbContext.PolymerGroups.Find(7), //PVC LSLTx
                _dbContext.PolymerGroups.Find(11) //PVC term
            };

            List<InsulatedBillet> kpsvvBillets;

            foreach(var group in polymerGroups)
            {
                kpsvvBillets = new List<InsulatedBillet>
                {
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor05,
                        Diameter = 1.8m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor075,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.06m : 1.96m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor1,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.29m : 2.19m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor15,
                        Diameter = 2.66m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor25,
                        Diameter = 3.06m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    }
                };
                _dbContext.AddRange(kpsvvBillets);
            }
            return _dbContext.SaveChanges();
        }
    }
}
