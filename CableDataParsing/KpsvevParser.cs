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
            var PVCLSLTxGroup = _dbContext.PolymerGroups.Find(7); //PVC LSLTx
            var PESelfExtinguish = _dbContext.PolymerGroups.Find(12); //PE self extinguish

            var polymerGroups = new List<PolymerGroup> { PVCGroup, PVCColdGroup, PVCTermGroup, PVCLSGroup };
            var polymerGroupsLTx = new List<PolymerGroup> { PVCLSLTxGroup };
            var polymerGroupsPE = new List<PolymerGroup> { PESelfExtinguish };
            var polymerGroupsFull = new List<PolymerGroup> { PVCGroup, PVCColdGroup, PVCTermGroup, PVCLSGroup, PESelfExtinguish };

            var TU2 = _dbContext.TechnicalConditions.Find(1);
            var TU30 = _dbContext.TechnicalConditions.Find(35);
            var TU49 = _dbContext.TechnicalConditions.Find(20);

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

            var twistedElementType = _dbContext.TwistedElementTypes.Find(2); //pair

            Cables.Common.CableProperty? cablePropSimple = null;
            Cables.Common.CableProperty? cablePropShield = Cables.Common.CableProperty.HasFoilShield;
            Cables.Common.CableProperty? cablePropKG = Cables.Common.CableProperty.HasArmourBraid;
            Cables.Common.CableProperty? cablePropKGShield = Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasArmourBraid;
            Cables.Common.CableProperty? cablePropK = Cables.Common.CableProperty.HasArmourBraid | Cables.Common.CableProperty.HasArmourTube;
            Cables.Common.CableProperty? cablePropKShield = Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasArmourBraid | Cables.Common.CableProperty.HasArmourTube;
            Cables.Common.CableProperty? cablePropB = Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube;
            Cables.Common.CableProperty? cablePropBShield = Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube;

            var configStart10 = new TableParserConfigurator(3, 2, 10, 5, 2, 1);
            var configStart8 = new TableParserConfigurator(3, 2, 8, 5, 2, 1);
            var configMid10 = new TableParserConfigurator(8, 2, 10, 5, 2, 1);
            var configMid8 = new TableParserConfigurator(8, 2, 8, 5, 2, 1);
            var configLTx1 = new TableParserConfigurator(3, 2, 9, 5, 2, 1);
            var configLTx2 = new TableParserConfigurator(9, 2, 9, 5, 2, 1);
            var configLTx3 = new TableParserConfigurator(15, 2, 9, 5, 2, 1);
            var configLTx4 = new TableParserConfigurator(21, 2, 9, 5, 2, 1);

            var configSet1 = (props: cablePropSimple, tu: TU2, configure: new List<(int tableIndex, TableParserConfigurator config)> { (0, configStart10), (1, configStart10) }, polymerGroups: polymerGroups);
            var configSet2 = (props: cablePropShield, tu: TU2, configure: new List<(int tableIndex, TableParserConfigurator config)> { (0, configMid10), (1, configMid8) }, polymerGroups: polymerGroups);
            var configSet3 = (props: cablePropSimple, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (3, configStart10), (4, configStart10) }, polymerGroups: polymerGroupsPE);
            var configSet4 = (props: cablePropShield, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (3, configMid10), (4, configMid8) }, polymerGroups: polymerGroupsPE);
            var configSet5 = (props: cablePropSimple, tu: TU49, configure: new List<(int tableIndex, TableParserConfigurator config)> { (2, configLTx1), (2, configLTx2) }, polymerGroups: polymerGroupsLTx);
            var configSet6 = (props: cablePropShield, tu: TU49, configure: new List<(int tableIndex, TableParserConfigurator config)> { (2, configLTx3), (2, configLTx4) }, polymerGroups: polymerGroupsLTx);
            var configSet7 = (props: cablePropKG, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (5, configStart10), (6, configStart10) }, polymerGroups: polymerGroupsFull);
            var configSet8 = (props: cablePropKGShield, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (5, configMid10), (6, configMid8) }, polymerGroups: polymerGroupsFull);
            var configSet9 = (props: cablePropK, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (7, configStart10), (8, configStart10) }, polymerGroups: polymerGroupsFull);
            var configSet10 = (props: cablePropKShield, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (7, configMid10), (8, configMid8) }, polymerGroups: polymerGroupsFull);
            var configSet11 = (props: cablePropB, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (9, configStart10) }, polymerGroups: polymerGroupsFull);
            var configSet12 = (props: cablePropBShield, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (9, configMid10) }, polymerGroups: polymerGroupsFull);

            var mainConfigArray = new[]
            { configSet1, configSet2, configSet3, configSet4, configSet5, configSet6, configSet7, configSet8, configSet9, configSet10, configSet11, configSet12 };

            var billets = _dbContext.InsulatedBillets.Where(b => b.CableShortNameId == 6)
                                                     .Include(p => p.Conductor)
                                                     .Include(p => p.PolymerGroup)
                                                     .ToList();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);

            var patternCable = new Cable
            {
                ClimaticMod = climaticModUHL,
                TwistedElementType = twistedElementType
            };

            var optionsCount = 1.0;

            foreach (var option in mainConfigArray)
            {
                var tableDataCommon = new List<TableCellData>();
                foreach (var set in option.configure)
                {
                    tableDataCommon.AddRange(_wordTableParser.GetCableCellsCollection(set.tableIndex, set.config));
                }

                foreach (var polymerGroup in option.polymerGroups)
                {
                    var insulationPolymerGroup = polymerGroup == PVCColdGroup || polymerGroup == PESelfExtinguish ? PVCGroup : polymerGroup;
                    var currentPolymerGroupBillets = (from b in billets
                                                      where b.PolymerGroup == insulationPolymerGroup
                                                      select b);
                    foreach (var tableCellData in tableDataCommon)
                    {
                        var cable = patternCable.Clone();
                        cable.CoverPolymerGroup = polymerGroup;
                        cable.TechnicalConditions = option.tu;
                        cable.CoverColor = polymerGroup == PVCColdGroup ? blackColor : redColor;
                        cable.OperatingVoltage = operatingVoltage;
                        cable.FireProtectionClass = PolymerGroupFireClassDict[polymerGroup];
                        ParseTableCellData(cable, tableCellData, currentPolymerGroupBillets, option.props, _splitter);
                        recordsCount++;
                    }
                }
                OnParseReport(optionsCount / mainConfigArray.Length);
                optionsCount++;
            }
            _wordTableParser.CloseWordApp();
            return recordsCount;
        }
    }
}
