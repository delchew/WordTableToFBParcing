using CableDataParsing.MSWordTableParsers;
using System;
using System.Collections.Generic;
using System.IO;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using CableDataParsing.CableTitleBulders;

namespace CableDataParsing
{
    public class Kevv_KerspParser : CableParser
    {
        private const int tablesCount = 4;

        public Kevv_KerspParser(string connectionString, FileInfo mSWordFile)
            : base(connectionString, mSWordFile, new Kevv_KerspTitleBuilder()) { }

        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;
            List<InsulatedBillet> pvcBillets, rubberBillets;
            pvcBillets = _dbContext.InsulatedBillets.AsNoTracking()
                                                   .Include(b => b.Conductor)
                                                   .Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэв"))
                                                   .ToList();
            rubberBillets = _dbContext.InsulatedBillets.AsNoTracking()
                                                  .Include(b => b.Conductor)
                                                  .Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэрс"))
                                                  .ToList();

            _wordTableParser = new WordTableParser(_mSWordFile).SetDataRowsCount(7)
                                                               .SetColumnHeadersRowIndex(3)
                                                               .SetRowHeadersColumnIndex(3)
                                                               .SetDataStartColumnIndex(4);

            var kersParams = new List<(int fireId, int polymerId, int colorId)>
            {
                (23, 4, 3), (26, 5, 8)
            };

            var dataStartRowIndexes = new int[2] { 4, 11 };

            List<TableCellData> tableData;

            for (int i = 0; i < tablesCount; i++)
            {
                _wordTableParser.DataColumnsCount = i % 2 == 0 ? 7 : 9;

                foreach (var index in dataStartRowIndexes)
                {
                    _wordTableParser.DataStartRowIndex = index;
                    tableData = _wordTableParser.GetCableCellsCollection(i + 1); //добавляем 1, потому что таблицы в MSWord нумеруются с 1, а не с 0.
                    foreach (var tableCellData in tableData)
                    {
                        if (i < 3)
                        {
                            ParseTableCellData(tableCellData, pvcBillets, _dbContext, ref recordsCount, index, (8, 6, 9));
                        }
                        else
                        {
                            foreach (var param in kersParams)
                            {
                                ParseTableCellData(tableCellData, rubberBillets, _dbContext, ref recordsCount, index, param);
                            }
                        }
                    }
                    tableData.Clear();
                    //ParseReport(672, recordsCount); //672 марки в таблицах
                }
            }
            return recordsCount;
        }

        private void ParseTableCellData(TableCellData tableCellData, List<InsulatedBillet> currentBilletsList, CablesContext dbContext,
                                            ref int recordsCount, int index, (int fireId, int polymerId, int colorId) kevvParams)
        {
            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
            {
                var billet = (from b in currentBilletsList
                              where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm
                              select b).First();
                var kevvKersp = new Cable
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
                
                Cables.Common.CableProperty? cableProp;
                if (index == 11)
                    cableProp = Cables.Common.CableProperty.HasBraidShield;
                else cableProp = null;

                kevvKersp.Title = cableTitleBuilder.GetCableTitle(kevvKersp, billet, cableProp); 
                var cableRec = dbContext.Cables.Add(kevvKersp);
                dbContext.SaveChanges();

                dbContext.ListCableBillets.Add(new ListCableBillets { BilletId = billet.Id, CableId = cableRec.Entity.Id });

                if (index == 11) //обрабатывается часть таблицы для кабелей с экраном
                    dbContext.ListCableProperties.Add(new ListCableProperties { PropertyId = 4, CableId = cableRec.Entity.Id });

                dbContext.SaveChanges();

                recordsCount++;
            }
            else
                throw new Exception($"Не удалось распарсить ячейку таблицы!");
        }
    }
}
