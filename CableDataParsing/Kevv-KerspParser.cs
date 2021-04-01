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
            var pvcBillets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэв"))
                                                   .Include(b => b.Conductor)
                                                   .AsNoTracking()
                                                   .ToList();
            var rubberBillets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэрс"))
                                                  .Include(b => b.Conductor)
                                                  .AsNoTracking()
                                                  .ToList();

            _wordTableParser = new WordTableParser().SetDataRowsCount(7)
                                                    .SetColumnHeadersRowIndex(3)
                                                    .SetRowHeadersColumnIndex(3)
                                                    .SetDataStartColumnIndex(4);

            var kersParams = new List<(int fireId, int polymerId, int colorId)>
            {
                (23, 4, 3), (26, 5, 8)
            };

            var cableProps = new List<Cables.Common.CableProperty?>
            {
                null, Cables.Common.CableProperty.HasBraidShield
            };

            List<TableCellData> tableData;

            _wordTableParser.OpenWordDocument(_mSWordFile);

            for (int i = 0; i < tablesCount; i++)
            {
                _wordTableParser.DataColumnsCount = i % 2 == 0 ? 7 : 9;

                foreach (var prop in cableProps)
                {
                    _wordTableParser.DataStartRowIndex = prop.HasValue ? 11 : 4;

                    tableData = _wordTableParser.GetCableCellsCollection(i + 1); //добавляем 1, потому что таблицы в MSWord нумеруются с 1, а не с 0.

                    foreach (var tableCellData in tableData)
                    {
                        if (i < 3)
                        {
                            ParseTableCellData(tableCellData, pvcBillets, prop, (8, 6, 9));
                        }
                        else
                        {
                            foreach (var param in kersParams)
                                ParseTableCellData(tableCellData, rubberBillets, prop, param);
                        }
                    }
                    tableData.Clear();
                    OnParseReport(cableBrandsCount, _recordsCount);
                }
            }
            _wordTableParser.CloseWordApp();
            return _recordsCount;
        }

        private void ParseTableCellData(TableCellData tableCellData, List<InsulatedBillet> currentBilletsList,
                                            Cables.Common.CableProperty? prop, (int fireId, int polymerId, int colorId) kevvParams)
        {
            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
            {
                decimal? maxCoverDiameter;
                if (decimal.TryParse(tableCellData.CellData, out decimal diameterValue))
                    maxCoverDiameter = diameterValue;
                else
                {
                    var cableSizes = tableCellData.CellData.Split('\u00D7'); //знак умножения в юникоде
                    if (cableSizes.Length < 2) return;
                    if (cableSizes.Length == 2 &&
                        decimal.TryParse(cableSizes[0], out decimal height) &&
                        decimal.TryParse(cableSizes[1], out decimal width))
                    {
                        maxCoverDiameter = null;
                    }
                    else throw new Exception("Wrong format table cell data!");
                }
                var billet = (from b in currentBilletsList
                              where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm
                              select b).First();
                var cable = new Cable
                {
                    ElementsCount = elementsCount,
                    TwistedElementTypeId = 1, //single
                    TechnicalConditionsId = 22, //ТУ 46
                    FireProtectionClassId = kevvParams.fireId,
                    CoverPolymerGroupId = kevvParams.polymerId,
                    CoverColorId = kevvParams.colorId,
                    MaxCoverDiameter = maxCoverDiameter,
                    ClimaticModId = 3, //УХЛ
                    OperatingVoltageId = 4
                };
                
                cable.Title = cableTitleBuilder.GetCableTitle(cable, billet, prop); 
                var cableRec = _dbContext.Cables.Add(cable).Entity;
                _dbContext.SaveChanges();

                _dbContext.ListCableBillets.Add(new ListCableBillets { BilletId = billet.Id, CableId = cableRec.Id });

                if (prop.HasValue)
                {
                    var listOfCableProperties = GetCableAssociatedPropertiesList(cableRec, prop.Value);
                    _dbContext.ListCableProperties.AddRange(listOfCableProperties);
                }

                _dbContext.SaveChanges();

                _recordsCount++;
            }
            else
                throw new Exception($"Не удалось распарсить ячейку таблицы!");
        }
    }
}
