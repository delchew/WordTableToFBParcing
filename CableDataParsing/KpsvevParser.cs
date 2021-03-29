using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CableDataParsing.MSWordTableParsers;
//using Cables.Common;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;
using WordObj = Microsoft.Office.Interop.Word;

namespace CableDataParsing
{
    public class KpsvevParser : CableParser
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public KpsvevParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
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
                    var PESelfExtinguish = _dbContext.PolymerGroups.Find(0); //PE self extinguish - add to DB!!! TODO

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

                    var cablePropertiesList = _dbContext.CableProperties.ToList();
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

                    _wordTableParser = new WordTableParser
                    {
                        ColumnHeadersRowIndex = 3,
                        RowHeadersColumnIndex = 2,
                        DataRowsCount = 5,
                        DataStartColumnIndex = 3,
                    };
                    List<TableCellData> tableData1, tableData2;

                    _wordTableParser.DataStartRowIndex = 4;
                    _wordTableParser.DataColumnsCount = 10;
                    tableData1 = _wordTableParser.GetCableCellsCollection(tables[1]);
                    _wordTableParser.DataColumnsCount = 8;
                    tableData2 = _wordTableParser.GetCableCellsCollection(tables[2]);
                    var tableDataNoShield = tableData1.Concat(tableData2);

                    foreach (var polymerGroup in polymerGroups)
                    {
                        foreach (var tableCellData in tableDataNoShield)
                        {
                            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                            decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                            decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                            {
                                var kpsvv = new Cable
                                {
                                    ClimaticMod = climaticModUHL,
                                    CoverPolymerGroup = polymerGroup,
                                    TechnicalConditions = TU2,
                                    CoverColor = redColor,
                                    OperatingVoltage = operatingVoltage,
                                    ElementsCount = elementsCount,
                                    MaxCoverDiameter = maxCoverDiameter,
                                    FireProtectionClass = PolymerGroupFireClassDict[polymerGroup],
                                    TwistedElementType = twistedElementType,
                                };
                                kpsvv.Title = GetCableTitle(kpsvv);
                                var cableRec = _dbContext.Cables.Add(kpsvv).Entity;
                                _dbContext.SaveChanges();
                            }
                            else
                            {
                                throw new NotImplementedException();
                            }
                        }
                    }
                    


                    _wordTableParser.DataStartRowIndex = 9;
                    _wordTableParser.DataColumnsCount = 10;
                    tableData1 = _wordTableParser.GetCableCellsCollection(tables[1]);
                    _wordTableParser.DataColumnsCount = 8;
                    tableData2 = _wordTableParser.GetCableCellsCollection(tables[2]);

                    var tableDataShield = tableData1.Concat(tableData2);

                    

                    

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

        public override string GetCableTitle(Cable cable, params object[] cableParametres)
        {
            var cableProps = cableParametres[0] as Cables.Common.CableProperty;

            _nameBuilder.Clear();
            _nameBuilder.Append("КПСВ");
            if ((cableProps & Cables.Common.CableProperty.HasFoilShield) == Cables.Common.CableProperty.HasFoilShield)
                _nameBuilder.Append("Э");
            string namePart = string.Empty;
            switch (cable.CoverPolymerGroup.Title)
            {
                case "PVC":
                case "PVC Term":
                case "PVC LS":
                case "PVC Cold":
                    namePart = "В";
                    break;
                case "PE":
                    namePart = "Пс";
                    break;
            }
            _nameBuilder.Append(namePart);

            if ((cableProps & Cables.Common.CableProperty.HasArmourBraid) == Cables.Common.CableProperty.HasArmourBraid)
            {
                if ((cableProps & Cables.Common.CableProperty.HasArmourTube) == Cables.Common.CableProperty.HasArmourTube)
                    _nameBuilder.Append($"К{namePart}");
                else
                    _nameBuilder.Append("КГ");
            }
            if ((cableProps & Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube) ==
                (Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube))
            {
                _nameBuilder.Append($"Б{namePart}");
            }

            if (cable.CoverPolymerGroup.Title == "PVC Term")
                _nameBuilder.Append("т");
            if (cable.CoverPolymerGroup.Title == "PVC Cold")
                _nameBuilder.Append("м");
            if (cable.CoverPolymerGroup.Title == "PVC LS")
                _nameBuilder.Append("нг(А)-LS");

            var cableConductorArea = cable.ListCableBillets.First().Billet.Conductor.AreaInSqrMm;
            namePart = Cables.Common.CableCalculations.FormatConductorArea((double)cableConductorArea);
            _nameBuilder.Append($" {cable.ElementsCount}х2х{namePart}");

            return _nameBuilder.ToString();
        }
    }
}
