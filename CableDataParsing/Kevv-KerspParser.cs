using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using CablesDatabaseEFCoreFirebird.Entities;
using CableDataParsing.CableTitleBulders;
using CableDataParsing.MSWordTableParsers;

namespace CableDataParsing
{
    public class Kevv_KerspParser : CableParser
    {
        private const int tablesCount = 4; // количество таблиц в документе
        private const int cableBrandsCount = 672;  //672 марки в таблицах
        private int _recordsCount;

        public Kevv_KerspParser(string connectionString, FileInfo mSWordFile)
            : base(connectionString, mSWordFile, new Kevv_KerspTitleBuilder()) { }

        public override int ParseDataToDatabase()
        {
            _recordsCount = 0;

            var twistedElementType = _dbContext.TwistedElementTypes.Where(t => t.ElementType.ToLower() =="single").First();
            var techCond = _dbContext.TechnicalConditions.Where(c => c.Title == "ТУ 16.К99-046-2011").First();
            var climaticMod = _dbContext.ClimaticMods.Where(m => m.Title == "УХЛ").First();
            var operatingVoltage = _dbContext.OperatingVoltages.Find(4);

            var fireClassPVCLS = _dbContext.FireProtectionClasses.Find(8);
            var fireClassHF = _dbContext.FireProtectionClasses.Find(23);
            var fireClassPURHF = _dbContext.FireProtectionClasses.Find(26);

            var polymerPVCLS = _dbContext.PolymerGroups.Find(6);
            var polymerHF = _dbContext.PolymerGroups.Find(4);
            var polymerPUR = _dbContext.PolymerGroups.Find(5);

            var colorGrey = _dbContext.Colors.Find(9);
            var colorBlack = _dbContext.Colors.Find(2);
            var colorOrange = _dbContext.Colors.Find(8);

            var pvcBillets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэв"))
                                                   .Include(b => b.Conductor)
                                                   .ToList();
            var rubberBillets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэрс"))
                                                  .Include(b => b.Conductor)
                                                  .ToList();

            _wordTableParser = new XceedWordTableParser().SetDataRowsCount(7)
                                                    .SetColumnHeadersRowIndex(3)
                                                    .SetRowHeadersColumnIndex(3)
                                                    .SetDataStartColumnIndex(4);

            var kersParams = new List<(FireProtectionClass fireClass, PolymerGroup polymer, Color color)>
            {
                (fireClassPVCLS, polymerPVCLS, colorGrey), (fireClassHF, polymerHF, colorBlack), (fireClassPURHF, polymerPUR, colorOrange)
            };

            var cableProps = new List<Cables.Common.CableProperty?>
            {
                null, Cables.Common.CableProperty.HasBraidShield
            };

            List<TableCellData> tableData;

            _wordTableParser.OpenWordDocument(_mSWordFile);

            for (int i = 0; i < tablesCount; i++)
            {
                _wordTableParser.DataColumnsCount = (i + 1) % 2 == 0 ? 7 : 9; //выбираем число столбцов в зависимости от чётности номера таблицы

                foreach (var prop in cableProps)
                {
                    _wordTableParser.DataStartRowIndex = prop.HasValue ? 11 : 4;

                    tableData = _wordTableParser.GetCableCellsCollection(i + 1); //добавляем 1, потому что таблицы в MSWord нумеруются с 1, а не с 0.

                    foreach (var tableCellData in tableData)
                    {
                        var cable = new Cable
                        {
                            TwistedElementType = twistedElementType,
                            TechnicalConditions = techCond,
                            ClimaticMod = climaticMod, 
                            OperatingVoltage = operatingVoltage
                        };
                        if (i < 2) //первые 2 таблицы для КЭВВ, остальные - КЭРс
                        {
                            cable.FireProtectionClass = fireClassPVCLS;
                            cable.CoverPolymerGroup = polymerPVCLS;
                            cable.CoverColor = colorGrey;
                            ParseTableCellData(cable, tableCellData, pvcBillets, prop);
                        }
                        else
                        {
                            foreach (var param in kersParams)
                            {
                                cable.FireProtectionClass = param.fireClass;
                                cable.CoverPolymerGroup = param.polymer;
                                cable.CoverColor = param.color;
                                ParseTableCellData(cable, tableCellData, rubberBillets, prop);
                            }
                        }
                    }
                    tableData.Clear();
                    OnParseReport(cableBrandsCount, _recordsCount);
                }
            }
            _wordTableParser.CloseWordApp();
            return _recordsCount;
        }

        private void ParseTableCellData(Cable cable, TableCellData tableCellData, List<InsulatedBillet> currentBilletsList,
                                            Cables.Common.CableProperty? cableProps, char splitter = ' ')
        {
            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
            {
                decimal height = 0m;
                decimal width = 0m;
                decimal? maxCoverDiameter;
                if (decimal.TryParse(tableCellData.CellData, out decimal diameterValue))
                    maxCoverDiameter = diameterValue;
                else
                {
                    var cableSizes = tableCellData.CellData.Split(splitter);
                    if (cableSizes.Length < 2) return;
                    if (cableSizes.Length == 2 &&
                        decimal.TryParse(cableSizes[0], out height) &&
                        decimal.TryParse(cableSizes[1], out width))
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
                cable.Title = cableTitleBuilder.GetCableTitle(cable, billet, cableProps);

                var cableRec = _dbContext.Cables.Add(cable).Entity;

                _dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet, Cable = cableRec });

                if (cableProps.HasValue)
                {
                    var listOfCableProperties = GetCableAssociatedPropertiesList(cableRec, cableProps.Value);
                    _dbContext.ListCableProperties.AddRange(listOfCableProperties);
                }
                if (!maxCoverDiameter.HasValue)
                {
                    var flatSize = new FlatCableSize { Height = height, Width = width, Cable = cableRec };
                    _dbContext.FlatCableSizes.Add(flatSize);
                }
                _dbContext.SaveChanges();

                _recordsCount++;
            }
            else
                throw new Exception($"Не удалось распарсить ячейку таблицы!");
        }
    }
}
