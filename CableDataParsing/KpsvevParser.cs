using System.Collections.Generic;
using System.IO;
using System.Linq;
using Cables.Common;
using CableDataParsing.CableBulders;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;
using System.Globalization;
using System;
using FirebirdDatabaseProvider;
using CableDataParsing.TableEntityes;

namespace CableDataParsing
{
    public class KpsvevParser : CableParser
    {
        private const char _splitter = '\u00D7'; //знак умножения в юникоде
        private KpsvevTitleBuilder _cableTitleBuilder;
        private FirebirdDBProvider _provider;
        private FirebirdDBTableProvider<CablePresenter> _cableTableProvider;
        private FirebirdDBTableProvider<ListCableBilletsPresenter> _listCableBilletsPresenter;
        private FirebirdDBTableProvider<ListCablePropertiesPresenter> _listCablePropertiesPresenter;
        private FirebirdDBTableProvider<FlatCableSizePresenter> _FlatCableSizePresenter;
        public KpsvevParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _cableTitleBuilder = new KpsvevTitleBuilder();
            _provider = new FirebirdDBProvider(connectionString);
            _cableTableProvider = new FirebirdDBTableProvider<CablePresenter>(_provider);
            _listCableBilletsPresenter = new FirebirdDBTableProvider<ListCableBilletsPresenter>(_provider);
            _listCablePropertiesPresenter = new FirebirdDBTableProvider<ListCablePropertiesPresenter>(_provider);
            _FlatCableSizePresenter = new FirebirdDBTableProvider<FlatCableSizePresenter>(_provider);
        }

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

            CablePropertySet? cablePropSimple = null;
            CablePropertySet? cablePropShield = CablePropertySet.HasFoilShield;
            CablePropertySet? cablePropKG = CablePropertySet.HasArmourBraid;
            CablePropertySet? cablePropKGShield = CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid;
            CablePropertySet? cablePropK = CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube;
            CablePropertySet? cablePropKShield = CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube;
            CablePropertySet? cablePropB = CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube;
            CablePropertySet? cablePropBShield = CablePropertySet.HasFoilShield | CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube;

            var configStart10 = new TableParserConfigurator(3, 2, 10, 5, 2, 1);
            var configStart8 = new TableParserConfigurator(3, 2, 8, 5, 2, 1);
            var configMid10 = new TableParserConfigurator(8, 2, 10, 5, 2, 1);
            var configMid8 = new TableParserConfigurator(8, 2, 8, 5, 2, 1);
            var configLTx1 = new TableParserConfigurator(3, 2, 9, 5, 2, 1);
            var configLTx2 = new TableParserConfigurator(9, 2, 9, 5, 8, 1);
            var configLTx3 = new TableParserConfigurator(15, 2, 9, 5, 14, 1);
            var configLTx4 = new TableParserConfigurator(21, 2, 9, 5, 20, 1);

            var configSet1 = (props: cablePropSimple, tu: TU2, configure: new List<(int tableIndex, TableParserConfigurator config)> { (0, configStart10), (1, configStart8) }, polymerGroups: polymerGroups);
            var configSet2 = (props: cablePropShield, tu: TU2, configure: new List<(int tableIndex, TableParserConfigurator config)> { (0, configMid10), (1, configMid8) }, polymerGroups: polymerGroups);
            var configSet3 = (props: cablePropSimple, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (3, configStart10), (4, configStart8) }, polymerGroups: polymerGroupsPE);
            var configSet4 = (props: cablePropShield, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (3, configMid10), (4, configMid8) }, polymerGroups: polymerGroupsPE);
            var configSet5 = (props: cablePropSimple, tu: TU49, configure: new List<(int tableIndex, TableParserConfigurator config)> { (2, configLTx1), (2, configLTx2) }, polymerGroups: polymerGroupsLTx);
            var configSet6 = (props: cablePropShield, tu: TU49, configure: new List<(int tableIndex, TableParserConfigurator config)> { (2, configLTx3), (2, configLTx4) }, polymerGroups: polymerGroupsLTx);
            var configSet7 = (props: cablePropKG, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (5, configStart10), (6, configStart8) }, polymerGroups: polymerGroupsFull);
            var configSet8 = (props: cablePropKGShield, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (5, configMid10), (6, configMid8) }, polymerGroups: polymerGroupsFull);
            var configSet9 = (props: cablePropK, tu: TU30, configure: new List<(int tableIndex, TableParserConfigurator config)> { (7, configStart10), (8, configStart8) }, polymerGroups: polymerGroupsFull);
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

            var cablePresenter = new CablePresenter { ClimaticModId = climaticModUHL.Id, TwistedElementTypeId = twistedElementType.Id };

            var optionsCount = 1.0;
            _provider.OpenConnection();
            try
            {
                foreach (var option in mainConfigArray)
                {
                    var tableDataCommon = new List<TableCellData>();
                    foreach (var (tableIndex, config) in option.configure)
                    {
                        tableDataCommon.AddRange(_wordTableParser.GetCableCellsCollection(tableIndex, config));
                    }

                    foreach (var polymerGroup in option.polymerGroups)
                    {
                        var insulationPolymerGroup = polymerGroup == PVCColdGroup || polymerGroup == PESelfExtinguish ? PVCGroup : polymerGroup;
                        var currentPolymerGroupBillets = (from b in billets
                                                          where b.PolymerGroup == insulationPolymerGroup
                                                          select b);
                        foreach (var tableCellData in tableDataCommon)
                        {
                            cablePresenter.CoverPolimerGroupId = polymerGroup.Id;
                            cablePresenter.TechCondId = option.tu.Id;
                            cablePresenter.CoverColorId = polymerGroup == PVCColdGroup ? blackColor.Id : redColor.Id;
                            cablePresenter.OperatingVoltageId = operatingVoltage.Id;
                            cablePresenter.FireProtectionId = PolymerGroupFireClassDict[polymerGroup].Id;
                            ParseTableCellData(cablePresenter, tableCellData, currentPolymerGroupBillets, option.props, _splitter);
                            recordsCount++;
                        }
                    }
                    OnParseReport(optionsCount / mainConfigArray.Length);
                    optionsCount++;
                }
            }
            finally
            {
                _wordTableParser.CloseWordApp();
                _provider.CloseConnection();
            }
            return recordsCount;
        }

        private void ParseTableCellData(CablePresenter cable, TableCellData tableCellData, IEnumerable<InsulatedBillet> currentBilletsList,
                                          CablePropertySet? cableProps = null, char splitter = ' ')
        {
            if (decimal.TryParse(tableCellData.ColumnHeaderData, NumberStyles.Any, _cultureInfo, out decimal elementsCount) &&
                decimal.TryParse(tableCellData.RowHeaderData, NumberStyles.Any, _cultureInfo, out decimal conductorAreaInSqrMm))
            {
                decimal height = 0m;
                decimal width = 0m;
                decimal? maxCoverDiameter;
                if (decimal.TryParse(tableCellData.CellData, NumberStyles.Any, _cultureInfo, out decimal diameterValue))
                    maxCoverDiameter = diameterValue;
                else
                {
                    var cableSizes = tableCellData.CellData.Split(splitter);
                    if (cableSizes.Length < 2) return;
                    if (cableSizes.Length == 2 &&
                        decimal.TryParse(cableSizes[0], NumberStyles.Any, _cultureInfo, out height) &&
                        decimal.TryParse(cableSizes[1], NumberStyles.Any, _cultureInfo, out width))
                    {
                        maxCoverDiameter = null;
                    }
                    else throw new Exception("Wrong format table cell data!");
                }
                var billet = (from b in currentBilletsList
                              where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm
                              select b).First();
                cable.ElementsCount = elementsCount;
                cable.MaxCoverDiameter = maxCoverDiameter;
                cable.Title = _cableTitleBuilder.GetCableTitle(cable, billet, cableProps);  //TODO!!! Make NameBuider!

                var cablePresenterId = _cableTableProvider.AddItem(cable);

                _listCableBilletsPresenter.AddItem(new ListCableBilletsPresenter { CableId = cablePresenterId, BilletId = billet.Id });

                var intProp = 0b_0000000001;

                for (int j = 0; j < cablePropertiesCount; j++)
                {
                    var Prop = (CablePropertySet)intProp;

                    if ((cableProps & Prop) == Prop)
                    {
                        var propertyObj = cablePropertiesList.Where(p => p.BitNumber == intProp).First();
                        _listCablePropertiesPresenter.AddItem(new ListCablePropertiesPresenter { PropertyId = propertyObj.Id, CableId = cablePresenterId });

                    }
                    intProp <<= 1;
                }

                if (!maxCoverDiameter.HasValue)
                {
                    _FlatCableSizePresenter.AddItem(new FlatCableSizePresenter { CableId = cablePresenterId, Height = height, Width = width });
                }
            }
            else
                throw new Exception($"Не удалось распарсить ячейку таблицы!");
        }
    }
}
