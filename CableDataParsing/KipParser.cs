using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using WordObj = Microsoft.Office.Interop.Word;


namespace CableDataParsing
{
    public class KipParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;

        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _connectionString;
        private StringBuilder _nameBuilder;

        public KipParser(string connectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _connectionString = connectionString;
        }

        public int ParseDataToDatabase()
        {
            int recordsCount = 0;
            ParseKipBillets();

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
                        DataRowsCount = 4,
                        DataColumnsCount = 11,
                        ColumnHeadersRowIndex = 3,
                        RowHeadersColumnIndex = 1, //здесь не будут учитываться
                        DataStartColumnIndex = 2
                    };
                    List<TableCellData> tableData;
                    using (var dbContext = new CablesContext(_connectionString))
                    {
                        var climaticMod = dbContext.ClimaticMods.Where(c => c.Title.ToUpper() == "").Single(); //написать условие
                        var blackColor = dbContext.Colors.Where(c => c.Title.ToLower() == "black").Single();
                        var greyColor = dbContext.Colors.Where(c => c.Title.ToLower() == "grey").Single();
                        var operatingVoltage = dbContext.OperatingVoltages.Where(o => o.ACVoltage == 300 && o.DCVoltage == null).Single();

                        var coverPolymerGroupList8TC = new List<PolymerGroup>
                        {
                            dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC").Single(),
                            dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PE").Single(),
                            dbContext.PolymerGroups.Add(new PolymerGroup{Title = "PVC Term", TitleRus = "ПВХ пластикат термостойкий"}).Entity,
                            dbContext.PolymerGroups.Add(new PolymerGroup{Title = "PVC Cold", TitleRus = "ПВХ пластикат морозостойкий"}).Entity
                        };

                        var hfPolymerGroup = dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "HFCOMPOUND").Single();

                        var coverPolymerGroupList25TC = new List<PolymerGroup>
                        {
                            dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC LS").Single(),
                            hfPolymerGroup
                        };
                        var coverPolymerGroupList42TC = new List<PolymerGroup>
                        {
                            hfPolymerGroup
                        };

                        var techCondPOlymerGroupsList = new List<(TechnicalConditions techCond, List<PolymerGroup> polymerGroups)>
                        {
                            (dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("008-2001")).Single(), coverPolymerGroupList8TC),
                            (dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("025-2005")).Single(), coverPolymerGroupList25TC),
                            (dbContext.TechnicalConditions.Where(t => t.Title.ToUpper().Contains("042-2010")).Single(), coverPolymerGroupList42TC)
                        };

                        var dataStartRowIndexes = new int[2] { 4, 8 };

                        foreach (var techCondPolymerGroup in techCondPOlymerGroupsList)
                        {
                            foreach (var polymerGroup in techCondPolymerGroup.polymerGroups)
                            {
                                foreach (var index in dataStartRowIndexes)
                                {
                                    _wordTableParser.DataStartRowIndex = index;
                                    tableData = _wordTableParser.GetCableCellsCollection(tables[1]);
                                    foreach (var tableCellData in tableData)
                                    {
                                        if (int.TryParse(tableCellData.ColumnHeaderData, out int elementsCount) &&
                                            decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter))
                                        {
                                            var kip = new Cable //Доделать!
                                            {
                                                ClimaticMod = climaticMod,
                                                CoverPolymerGroup = polymerGroup,
                                                TechnicalConditions = techCondPolymerGroup.techCond,
                                                CoverColor = polymerGroup.Title == "PVC" || polymerGroup.Title == "PVC Term" || polymerGroup.Title == "PVC LS" ? greyColor : blackColor,
                                                OperatingVoltage = operatingVoltage,

                                            };
                                        }
                                        else
                                            continue;
                                    }
                                    tableData.Clear();
                                    
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
                app.Quit();
            }

            return recordsCount;
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
