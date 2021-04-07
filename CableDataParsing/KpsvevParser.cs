using System.Collections.Generic;
using System.IO;
using System.Linq;
using CableDataParsing.CableBulders;
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

            var PVCGroup = _dbContext.PolymerGroups.Find(1); // PVC
            var PVCColdGroup = _dbContext.PolymerGroups.Find(10); // PVC Cold
            var PVCTermGroup = _dbContext.PolymerGroups.Find(11); //PVC term
            var PVCLSGroup = _dbContext.PolymerGroups.Find(6); //PVC LS

            var polymerGroups = new List<PolymerGroup> { PVCGroup, PVCColdGroup, PVCTermGroup, PVCLSGroup };

            var PVCLSLTxGroup = _dbContext.PolymerGroups.Find(7); //PVC LSLTx
            var PESelfExtinguish = _dbContext.PolymerGroups.Find(12); //PE self extinguish

            var noFireProtectClass = _dbContext.FireProtectionClasses.Find(1); //О1.8.2.5.4
            var LSFireProtectClass = _dbContext.FireProtectionClasses.Find(8); //П1б.8.2.2.2
            var LSLTxFireProtectClass = _dbContext.FireProtectionClasses.Find(28); //П1б.8.2.1.2

            var PolymerGroupFireClassDict = new Dictionary<PolymerGroup, FireProtectionClass>()
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

            _wordTableParser = new XceedWordTableParser().SetColumnHeadersRowIndex(2)
                                                         .SetRowHeadersColumnIndex(1)
                                                         .SetDataRowsCount(5)
                                                         .SetDataStartColumnIndex(2);

            IEnumerable<TableCellData> tableData1, tableData2;

            _wordTableParser.OpenWordDocument(_mSWordFile);

            var patternCable = new Cable
            {
                ClimaticMod = climaticModUHL,
                TwistedElementType = twistedElementType
            };
            foreach (var cableProps in kpsvevCableProps)
            {
                _wordTableParser.DataStartRowIndex = cableProps.HasValue ? 8 : 3;
                _wordTableParser.DataColumnsCount = 10;
                tableData1 = _wordTableParser.GetCableCellsCollection(0);
                _wordTableParser.DataColumnsCount = 8;
                tableData2 = _wordTableParser.GetCableCellsCollection(1);
                var tableDataCommon = tableData1.Concat(tableData2);

                foreach (var polymerGroup in polymerGroups)
                {
                    var insulationPolymerGroup = polymerGroup == PVCColdGroup ? PVCGroup : polymerGroup;
                    var currentPolymerGroupBillets = (from b in billets
                                                      where b.PolymerGroup == insulationPolymerGroup
                                                      select b);
                    foreach(var tableCellData in tableDataCommon)
                    {
                        var cable = patternCable.Clone();
                        cable.CoverPolymerGroup = polymerGroup;
                        cable.TechnicalConditions = TU2;
                        cable.CoverColor = polymerGroup == PVCColdGroup ? blackColor : redColor;
                        cable.OperatingVoltage = operatingVoltage;
                        cable.FireProtectionClass = PolymerGroupFireClassDict[polymerGroup];
                        ParseTableCellData(cable, tableCellData, currentPolymerGroupBillets, cableProps, _splitter);
                        recordsCount++;
                    }
                    //OnParseReport((double)_recordsCount / CABLE_BRANDS_COUNT);
                }
            }
            _wordTableParser.CloseWordApp();
            return recordsCount;
        }
    }
}
