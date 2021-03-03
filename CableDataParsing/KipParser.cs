using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;
using WordObj = Microsoft.Office.Interop.Word;


namespace CableDataParsing
{
    public class KipParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;

        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _connectionString;
        private StringBuilder _nameBuilder = new StringBuilder();
        private int cablePropertiesCount = Enum.GetNames(typeof(Cables.Common.CableProperty)).Count();
        private const int POLYMER_GROUP_COUNT = 7;

        public KipParser(string connectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _connectionString = connectionString;
        }

        public int ParseDataToDatabase()
        {
            int recordsCount = 0;
            //ParseKipBillets();

            var app = new WordObj.Application { Visible = false };
            object fileName = _mSWordFile.FullName;

            try
            {
                app.Documents.Open(ref fileName);
                var document = app.ActiveDocument;
                var tables = document.Tables;

                if (tables.Count > 0)
                {
                    _wordTableParser = new WordTableParser
                    {
                        DataRowsCount = 1,
                        DataColumnsCount = 11,
                        ColumnHeadersRowIndex = 3,
                        DataStartColumnIndex = 2
                    };
                    List<TableCellData> tableData;
                    using (var dbContext = new CablesContext(_connectionString))
                    {
                        var cablePropertiesList = dbContext.CableProperties.ToList();

                        var climaticModV = dbContext.ClimaticMods.Where(c => c.Id == 7).Single();
                        var climaticModUHL = dbContext.ClimaticMods.Where(c => c.Id == 3).Single();

                        var blackColor = dbContext.Colors.Where(c => c.Title.ToLower() == "black").Single();
                        var greyColor = dbContext.Colors.Where(c => c.Title.ToLower() == "grey").Single();
                        var operatingVoltage = dbContext.OperatingVoltages.Where(o => o.ACVoltage == 300 && o.DCVoltage == null).Single();

                        var noFireClass = dbContext.FireProtectionClasses.Where(f => f.Id == 1).Single();
                        var lsFireClass = dbContext.FireProtectionClasses.Where(f => f.Id == 8).Single();
                        var hfFireClass = dbContext.FireProtectionClasses.Where(f => f.Id == 13).Single();

                        var cond078 = dbContext.Conductors.Where(c => c.WiresDiameter == 0.26m && c.WiresCount == 7 && c.MetalId == 2).Single();
                        var cond060 = dbContext.Conductors.Where(c => c.WiresDiameter == 0.20m && c.WiresCount == 7 && c.MetalId == 2).Single();
                        var cableShortName = dbContext.CableShortNames.Where(n => n.ShortName == "КИП").Single();
                        var twistedElementType = dbContext.TwistedElementTypes.Where(t => t.Id == 2).Single();

                        var billet060_150 = dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond060.Id && b.Diameter == 1.50m && b.CableShortName.ShortName == cableShortName.ShortName).Single();
                        var billet060_142 = dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond060.Id && b.Diameter == 1.42m && b.CableShortName.ShortName == cableShortName.ShortName).Single();
                        var billet078PE = dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond078.Id && b.Diameter == 1.78m && b.CableShortName.ShortName == cableShortName.ShortName).Single();
                        var billet078PVC = dbContext.InsulatedBillets.Include(b => b.Conductor).Where(b => b.ConductorId == cond078.Id && b.Diameter == 1.60m && b.CableShortName.ShortName == cableShortName.ShortName).Single();

                        var coverPolymerGroupList8TC = new List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)>
                        {
                            (dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC").Single(), noFireClass),
                            (dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PE").Single(), noFireClass),
                            (dbContext.PolymerGroups.Add(new PolymerGroup{Title = "PVC Term", TitleRus = "ПВХ пластикат термостойкий"}).Entity, noFireClass),
                            (dbContext.PolymerGroups.Add(new PolymerGroup{Title = "PVC Cold", TitleRus = "ПВХ пластикат морозостойкий"}).Entity, noFireClass)
                        };

                        var hfPolymerGroup = dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "HFCOMPOUND").Single();

                        var coverPolymerGroupList25TC = new List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)>
                        {
                            (dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC LS").Single(), lsFireClass),
                            (hfPolymerGroup, hfFireClass)
                        };
                        var coverPolymerGroupList42TC = new List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)>
                        {
                            (hfPolymerGroup, hfFireClass)
                        };

                        var techCondPolymerGroupsList = new List<(TechnicalConditions techCond, List<(PolymerGroup polymerGroup, FireProtectionClass fireClass)> paramGroups)>
                        {
                            (dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("008-2001")).Single(), coverPolymerGroupList8TC),
                            (dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("025-2005")).Single(), coverPolymerGroupList25TC),
                            (dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("042-2010")).Single(), coverPolymerGroupList42TC)
                        };

                        var cableProps = new List<Cables.Common.CableProperty>
                        {
                            Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasBraidShield,
                            Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasBraidShield | Cables.Common.CableProperty.HasArmourBraid,
                            Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasBraidShield | Cables.Common.CableProperty.HasArmourBraid | Cables.Common.CableProperty.HasArmourTube,
                            Cables.Common.CableProperty.HasFoilShield | Cables.Common.CableProperty.HasBraidShield | Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube
                        };

                        var dataStartRowIndexes = new int[2] { 4, 8 };
                        InsulatedBillet billet;
                        var flagBG = false;
                        foreach (var techCondPolymerGroup in techCondPolymerGroupsList) //Не оптимально, но для разового метода сойдёт, если отделять БГ кабель от остальных со своей таблицей, придётся больше логики писать, пусть лучше чуть дольше работает метод)))
                        {
                            var tableNumber = techCondPolymerGroup.paramGroups == coverPolymerGroupList42TC ? 2 : 1;
                            if (tableNumber == 2) flagBG = true;
                            var currentPolymerGroupCount = 1;
                            foreach (var paramGroup in techCondPolymerGroup.paramGroups)
                            {
                                foreach (var index in dataStartRowIndexes)
                                {
                                    _wordTableParser.DataStartRowIndex = index;

                                    foreach (var cableProp in cableProps)
                                    {
                                        tableData = _wordTableParser.GetCableCellsCollection(tables[tableNumber]);
                                        foreach (var tableCellData in tableData)
                                        {
                                            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                                                decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter))
                                            {
                                                if (index < 8)
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
                                                    Title = GetKipName(elementsCount, paramGroup.polymerGroup, billet, cableProp, flagBG)
                                                };
                                                var cableRec = dbContext.Cables.Add(kip).Entity;
                                                dbContext.SaveChanges();

                                                dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet, Cable = cableRec });
                                                if (elementsCount == 1.5m)
                                                    dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet078PVC, Cable = cableRec });

                                                var intProp = 0b_0000000001;
                                                for (int i = 0; i < cablePropertiesCount; i++)
                                                {
                                                    var Prop = (Cables.Common.CableProperty)intProp;

                                                    if ((cableProp & Prop) == Prop)
                                                    {
                                                        var propertyObj = cablePropertiesList.Where(p => p.BitNumber == (int)Prop).Single();
                                                        dbContext.ListCableProperties.Add(new ListCableProperties { Property = propertyObj, Cable = cableRec });
                                                    }
                                                    intProp <<= 1;
                                                }
                                                dbContext.SaveChanges();

                                                recordsCount++;
                                            }
                                            else
                                                continue;
                                        }
                                        tableData.Clear();
                                        _wordTableParser.DataStartRowIndex++;
                                    }
                                }
                                ParseReport(POLYMER_GROUP_COUNT, currentPolymerGroupCount);
                                currentPolymerGroupCount++;
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
                app.Quit();
            }

            return recordsCount;
        }

        private string GetKipName(decimal elementsCount, PolymerGroup polymerGroup, InsulatedBillet billet, Cables.Common.CableProperty cableProp, bool flagBG)
        {
            _nameBuilder.Clear();
            _nameBuilder.Append("КИП");
            string namePart, condDiam;
            if (billet.Conductor.ConductorDiameter == 0.60m)
            {
                namePart = "Э";
                condDiam = "0,60";
            }
            else
            {
                namePart = "вЭ";
                condDiam = "0,78";
            }
            _nameBuilder.Append(namePart);
            switch (polymerGroup.Title)
            {
                case "PVC":
                case "PVC Term":
                case "PVC LS":
                case "PVC Cold": namePart = "В";
                    break;
                case "PE": namePart = "П";
                    break;
                default: namePart = string.Empty;
                    break;
            }
            _nameBuilder.Append(namePart);

            if ((cableProp & Cables.Common.CableProperty.HasArmourBraid) == Cables.Common.CableProperty.HasArmourBraid)
            {
                if ((cableProp & Cables.Common.CableProperty.HasArmourTube) == Cables.Common.CableProperty.HasArmourTube)
                    _nameBuilder.Append($"К{namePart}");
                else
                    _nameBuilder.Append("КГ");
            }    
            if((cableProp & Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube) ==
                (Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube))
            {
                _nameBuilder.Append($"Б{namePart}");
            }

            if (polymerGroup.Title == "PVC Term")
                _nameBuilder.Append("т");
            if (polymerGroup.Title == "PVC Cold")
                _nameBuilder.Append("м");
            if (polymerGroup.Title == "PVC LS")
                _nameBuilder.Append("нг(А)-LS");
            if (polymerGroup.Title.ToUpper() == "HFCOMPOUND")
            {
                if (flagBG)
                    _nameBuilder.Append("нг(А)-БГ");
                else
                    _nameBuilder.Append("нг(А)-HF");
            }

            _nameBuilder.Append($" {elementsCount}х2х{condDiam}");

            return _nameBuilder.ToString();
        }

        public void ParseKipBillets()
        {
            using(var dbContext = new CablesContext(_connectionString))
            {
                var cond078 = dbContext.Conductors.Where(c => c.WiresDiameter == 0.26m && c.WiresCount == 7 && c.MetalId == 2).Single();
                var cond060 = dbContext.Conductors.Where(c => c.WiresDiameter == 0.20m && c.WiresCount == 7 && c.MetalId == 2).Single();
                var polymerGroupPE = dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PE").Single();
                var polymerGroupPVC = dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC").Single();

                var operatingVoltage = dbContext.OperatingVoltages.Add(new OperatingVoltage { ACVoltage = 300, ACFriquency = 50, Description = "Переменное - 300В, 50Гц" }).Entity;
                var cableShortName = dbContext.CableShortNames.Add(new CableShortName { ShortName = "КИП" }).Entity;
                var foamPE = dbContext.PolymerGroups.Add(new PolymerGroup { Title = "Foamed PE", TitleRus = "Вспененный полиэтилен" }).Entity;

                var kipBillets = new List<InsulatedBillet>
                {
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName, 
                        Conductor = cond060, 
                        Diameter = 1.42m, 
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPE
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond060,
                        Diameter = 1.50m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPE
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond078,
                        Diameter = 1.78m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = foamPE
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond078,
                        Diameter = 1.60m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPVC
                    }
                };

                dbContext.InsulatedBillets.AddRange(kipBillets);
                dbContext.SaveChanges();
            }
        }
    }
}
