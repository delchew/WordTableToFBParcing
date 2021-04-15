using System.Collections.Generic;
using System.IO;
using System.Linq;
using Cables.Common;
using CableDataParsing.CableBulders;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;

namespace CableDataParsing
{
    public class KipParser : CableParser
    {
        private const int POLYMER_GROUP_COUNT = 7;
        private KipTitleBuilder _cableTitleBuilder;

        public KipParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _cableTitleBuilder = new KipTitleBuilder();
        }

        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;
            var configurator = new TableParserConfigurator().SetDataRowsCount(1)
                                                         .SetDataColumnsCount(11)
                                                         .SetColumnHeadersRowIndex(2)
                                                         .SetDataStartColumnIndex(1);
            _wordTableParser = new XceedWordTableParser();

            IEnumerable<TableCellData> tableData;
            var cablePropertiesList = _dbContext.CableProperties.ToList();

            var climaticModV = _dbContext.ClimaticMods.Where(c => c.Id == 7).Single();
            var climaticModUHL = _dbContext.ClimaticMods.Where(c => c.Id == 3).Single();

            var blackColor = _dbContext.Colors.Where(c => c.Title.ToLower() == "black").Single();
            var greyColor = _dbContext.Colors.Where(c => c.Title.ToLower() == "grey").Single();
            var operatingVoltage = _dbContext.OperatingVoltages.Where(o => o.ACVoltage == 300 && o.DCVoltage == null).Single();

            var noFireClass = _dbContext.FireProtectionClasses.Where(f => f.Id == 1).Single();
            var lsFireClass = _dbContext.FireProtectionClasses.Where(f => f.Id == 8).Single();
            var hfFireClass = _dbContext.FireProtectionClasses.Where(f => f.Id == 13).Single();

            var cond078 = _dbContext.Conductors.Where(c => c.WiresDiameter == 0.26m && c.WiresCount == 7 && c.MetalId == 2).Single();
            var cond060 = _dbContext.Conductors.Where(c => c.WiresDiameter == 0.20m && c.WiresCount == 7 && c.MetalId == 2).Single();
            var cableShortName = _dbContext.CableShortNames.Where(n => n.ShortName == "КИП").Single();
            var twistedElementType = _dbContext.TwistedElementTypes.Where(t => t.Id == 2).Single();

            var billet060_150 = _dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond060.Id && b.Diameter == 1.50m && b.CableShortName.ShortName == cableShortName.ShortName).Single();
            var billet060_142 = _dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond060.Id && b.Diameter == 1.42m && b.CableShortName.ShortName == cableShortName.ShortName).Single();
            var billet078PE = _dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond078.Id && b.Diameter == 1.78m && b.CableShortName.ShortName == cableShortName.ShortName).Single();
            var billet078PVC = _dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond078.Id && b.Diameter == 1.60m && b.CableShortName.ShortName == cableShortName.ShortName).Single();

            var coverPolymerGroupList8TC = new List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)>
            {
                (_dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC").Single(), noFireClass),
                (_dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PE").Single(), noFireClass),
                (_dbContext.PolymerGroups.Add(new PolymerGroup{Title = "PVC Term", TitleRus = "ПВХ пластикат термостойкий"}).Entity, noFireClass),
                (_dbContext.PolymerGroups.Add(new PolymerGroup{Title = "PVC Cold", TitleRus = "ПВХ пластикат морозостойкий"}).Entity, noFireClass)
            };

            var hfPolymerGroup = _dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "HFCOMPOUND").Single();

            var coverPolymerGroupList25TC = new List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)>
            {
                (_dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC LS").Single(), lsFireClass),
                (hfPolymerGroup, hfFireClass)
            };
            var coverPolymerGroupList42TC = new List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)>
            {
                (hfPolymerGroup, hfFireClass)
            };

            var techCondPolymerGroupsList = new List<(TechnicalConditions techCond, List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)> paramGroups)>
            {
                (_dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("008-2001")).Single(), coverPolymerGroupList8TC),
                (_dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("025-2005")).Single(), coverPolymerGroupList25TC),
                (_dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("042-2010")).Single(), coverPolymerGroupList42TC)
            };

            var cableProps = new List<CablePropertySet>
            {
                CablePropertySet.HasFoilShield | CablePropertySet.HasBraidShield,
                CablePropertySet.HasFoilShield | CablePropertySet.HasBraidShield | CablePropertySet.HasArmourBraid,
                CablePropertySet.HasFoilShield | CablePropertySet.HasBraidShield | CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube,
                CablePropertySet.HasFoilShield | CablePropertySet.HasBraidShield | CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube
            };

            var dataStartRowIndexes = new int[2] { 3, 7 };
            InsulatedBillet billet;
            var flagBG = false;


            _wordTableParser.OpenWordDocument(_mSWordFile);
            var currentPolymerGroupCount = 1;
            foreach (var techCondPolymerGroup in techCondPolymerGroupsList) //Не оптимально, но для разового метода сойдёт, если отделять БГ кабель от остальных со своей таблицей, придётся больше логики писать, пусть лучше чуть дольше работает метод)))
            {
                var tableIndex = techCondPolymerGroup.paramGroups == coverPolymerGroupList42TC ? 1 : 0;
                
                if (tableIndex == 1) flagBG = true;
                
                foreach (var paramGroup in techCondPolymerGroup.paramGroups)
                {
                    foreach (var index in dataStartRowIndexes)
                    {
                        configurator.DataStartRowIndex = index;

                        foreach (var cableProp in cableProps)
                        {
                            tableData = _wordTableParser.GetCableCellsCollection(tableIndex, configurator);
                            foreach (var tableCellData in tableData)
                            {
                                if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                                    decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter))
                                {
                                    if (index < 7)
                                    {
                                        if (elementsCount < 4)
                                            billet = billet060_150;
                                        else
                                            billet = billet060_142;
                                    }
                                    else
                                        billet = billet078PE;

                                    var kip = new Cable
                                    {
                                        ClimaticMod = paramGroup.polymerGroup.Title == "PVC LS" ? climaticModUHL : climaticModV,
                                        CoverPolymerGroup = paramGroup.polymerGroup,
                                        TechnicalConditions = techCondPolymerGroup.techCond,
                                        CoverColor = paramGroup.polymerGroup.Title == "PVC" || paramGroup.polymerGroup.Title == "PVC Term" || paramGroup.polymerGroup.Title == "PVC LS" ? greyColor : blackColor,
                                        OperatingVoltage = operatingVoltage,
                                        ElementsCount = elementsCount,
                                        MaxCoverDiameter = maxCoverDiameter,
                                        FireProtectionClass = paramGroup.fireClass,
                                        TwistedElementType = twistedElementType,
                                    };

                                    kip.Title = _cableTitleBuilder.GetCableTitle(kip, billet, cableProp, flagBG);
                                    var cableRec = _dbContext.Cables.Add(kip).Entity;

                                    _dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet, Cable = cableRec });
                                    if (elementsCount == 1.5m)
                                        _dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet078PVC, Cable = cableRec });

                                    var propList = GetCableAssociatedPropertiesList(cableRec, cableProp);

                                    _dbContext.AddRange(propList);

                                    _dbContext.SaveChanges();

                                    recordsCount++;
                                }
                                else
                                    continue;
                            }
                            configurator.DataStartRowIndex++;
                        }
                    }
                    OnParseReport((double)currentPolymerGroupCount / POLYMER_GROUP_COUNT);
                    currentPolymerGroupCount++;
                }
            }
            _wordTableParser.CloseWordApp();
            return recordsCount;
        }
    }
}
