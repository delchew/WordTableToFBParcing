using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CableDataParsing.CableTitleBulders;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;

namespace CableDataParsing
{
    public class KpsvevParser : CableParser
    {
        private const char _splitter = '\u00D7'; //знак умножения в юникоде
        public KpsvevParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile, new KpsvevTitleBuilder())
        { }

        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;

            var PVCGroup = _dbContext.PolymerGroups.Find(1).Id; // PVC
            var PVCColdGroup = _dbContext.PolymerGroups.Find(10).Id; // PVC Cold
            var PVCTermGroup = _dbContext.PolymerGroups.Find(11).Id; //PVC term
            var PVCLSGroup = _dbContext.PolymerGroups.Find(6).Id; //PVC LS

            var polymerGroups = new List<int> { PVCGroup, PVCColdGroup, PVCTermGroup, PVCLSGroup };

            var PVCLSLTxGroup = _dbContext.PolymerGroups.Find(7).Id; //PVC LSLTx
            var PESelfExtinguish = _dbContext.PolymerGroups.Find(12).Id; //PE self extinguish

            var noFireProtectClass = _dbContext.FireProtectionClasses.Find(1); //О1.8.2.5.4
            var LSFireProtectClass = _dbContext.FireProtectionClasses.Find(8); //П1б.8.2.2.2
            var LSLTxFireProtectClass = _dbContext.FireProtectionClasses.Find(28); //П1б.8.2.1.2

            var PolymerGroupFireClassDict = new Dictionary<int, FireProtectionClass>()
            {
                { PVCGroup, noFireProtectClass },
                { PVCColdGroup, noFireProtectClass },
                { PVCTermGroup, noFireProtectClass },
                { PVCLSGroup, LSFireProtectClass },
                { PVCLSLTxGroup, LSLTxFireProtectClass },
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

            var kpsvevCableProps = new List<Cables.Common.CableProperty?> { null, Cables.Common.CableProperty.HasFoilShield };

            var billets = _dbContext.InsulatedBillets.Where(b => b.CableShortNameId == 6)
                                                     .Include(p => p.Conductor)
                                                     .Include(p => p.PolymerGroup)
                                                     .ToList();

            _wordTableParser = new MSWordTableParser().SetColumnHeadersRowIndex(3)
                                                    .SetRowHeadersColumnIndex(2)
                                                    .SetDataRowsCount(5)
                                                    .SetDataStartColumnIndex(3);

            List<TableCellData> tableData1, tableData2;

            decimal? maxCoverDiameter;
            foreach (var prop in kpsvevCableProps)
            {
                _wordTableParser.DataStartRowIndex = prop.HasValue ? 9 : 4;
                _wordTableParser.DataColumnsCount = 10;
                tableData1 = _wordTableParser.GetCableCellsCollection(1);
                _wordTableParser.DataColumnsCount = 8;
                tableData2 = _wordTableParser.GetCableCellsCollection(2);
                var tableDataNoShield = tableData1.Concat(tableData2);

                foreach (var tableCellData in tableDataNoShield)
                {
                    if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                        decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                    {
                        decimal height = 0m;
                        decimal width = 0m;
                        if (decimal.TryParse(tableCellData.CellData, out decimal diameterValue))
                            maxCoverDiameter = diameterValue;
                        else
                        {
                            var cableSizes = tableCellData.CellData.Split(_splitter);
                            if (cableSizes.Length < 2) continue;
                            if (cableSizes.Length == 2 &&
                                decimal.TryParse(cableSizes[0], out height) &&
                                decimal.TryParse(cableSizes[1], out width))
                            {
                                maxCoverDiameter = null;
                            }
                            else throw new Exception("Wrong format table cell data!");
                        }

                        foreach (var polymerGroupId in polymerGroups)
                        {
                            var billet = (from b in billets
                                          where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm &&
                                          b.PolymerGroup.Id == polymerGroupId
                                          select b).First();

                            var cable = new Cable
                            {
                                ClimaticModId = climaticModUHL.Id,
                                CoverPolymerGroupId = polymerGroupId,
                                TechnicalConditionsId = TU2.Id,
                                CoverColorId = polymerGroupId == PVCColdGroup ? blackColor.Id : redColor.Id,
                                OperatingVoltageId = operatingVoltage.Id,
                                ElementsCount = elementsCount,
                                MaxCoverDiameter = maxCoverDiameter,
                                FireProtectionClassId = PolymerGroupFireClassDict[polymerGroupId].Id,
                                TwistedElementTypeId = twistedElementType.Id,
                            };
                            var cableRec = _dbContext.Cables.Add(cable).Entity;
                            //_dbContext.SaveChanges();

                            _dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet, Cable = cableRec });
                            cableRec.Title = cableTitleBuilder.GetCableTitle(cableRec, billet, prop);
                            //_dbContext.SaveChanges();

                            if (prop.HasValue)
                            {
                                var listOfCableProperties = GetCableAssociatedPropertiesList(cableRec, prop.Value);
                                _dbContext.ListCableProperties.AddRange(listOfCableProperties);
                            }
                            if (!maxCoverDiameter.HasValue)
                            {
                                var flatSize = new FlatCableSize { Height = height, Width = width, Cable = cableRec };
                                _dbContext.FlatCableSizes.Add(flatSize);
                            }

                            _dbContext.SaveChanges();

                            recordsCount++;
                        }
                    }
                    else
                        throw new InvalidCastException("Can't cast elementsCount or conductorAreaInSqrMm!");
                }
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
