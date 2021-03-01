using CableDataParsing.MSWordTableParsers;
using System;
using System.Collections.Generic;
using System.IO;
using WordObj = Microsoft.Office.Interop.Word;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using System.Linq;
using Microsoft.EntityFrameworkCore;
using System.Text;
using Cables.Common;

namespace CableDataParsing
{
    public class Kevv_KerspParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;

        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _connectionString;
        private StringBuilder _nameBuilder;


        public Kevv_KerspParser(string connectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _connectionString = connectionString;
        }

        public int ParseDataToDatabase()
        {
            int recordsCount = 0;
            List<InsulatedBillet> pvcBillets, rubberBillets;
            using (var dbContext = new CablesContext(_connectionString))
            {
                pvcBillets = dbContext.InsulatedBillets.AsNoTracking()
                                                       .Include(b => b.Conductor)
                                                       .Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэв"))
                                                       .ToList();
                rubberBillets = dbContext.InsulatedBillets.AsNoTracking()
                                                          .Include(b => b.Conductor)
                                                          .Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэрс"))
                                                          .ToList();
            }

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
                        DataRowsCount = 7,
                        ColumnHeadersRowIndex = 3,
                        RowHeadersColumnIndex = 3,
                        DataStartColumnIndex = 4,
                    };

                    var kersParams = new List<(int fireId, int polymerId, int colorId)>
                    {
                        (23, 4, 3), (26, 5, 8)
                    };

                    var dataStartRowIndexes = new int[2] { 4, 11 };

                    List<TableCellData> tableData;

                    using (var dbContext = new CablesContext(_connectionString))
                    {
                        for (int i = 1; i <= tables.Count; i++)
                        {
                            _wordTableParser.DataColumnsCount = i % 2 == 0 ? 7 : 9;

                            foreach (var index in dataStartRowIndexes)
                            {
                                _wordTableParser.DataStartRowIndex = index;
                                tableData = _wordTableParser.GetCableCellsCollection(tables[i]);
                                foreach (var tableCellData in tableData)
                                {
                                    if (i < 3)
                                    {
                                        ParseTableCellData(tableCellData, pvcBillets, dbContext, ref recordsCount, index, (8, 6, 9));
                                    }
                                    else
                                    {
                                        foreach (var param in kersParams)
                                        {
                                            ParseTableCellData(tableCellData, rubberBillets, dbContext, ref recordsCount, index, param);
                                        }
                                    }
                                }
                                tableData.Clear();
                                ParseReport(672, recordsCount); //672 марки в таблицах
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

        private void ParseTableCellData(TableCellData tableCellData, List<InsulatedBillet> currentBilletsList, CablesContext dbContext, ref int recordsCount, int index, (int fireId, int polymerId, int colorId) kevvParams)
        {
            if (int.TryParse(tableCellData.ColumnHeaderData, out int elementsCount) &&
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
                kevvKersp.Title = BuildTitle(kevvKersp, billet, index == 11 ? "Э" : string.Empty);
                var cableRec = dbContext.Cables.Add(kevvKersp);
                dbContext.SaveChanges();

                dbContext.ListCableBillets.Add(new ListCableBillets { BilletId = billet.Id, CableId = cableRec.Entity.Id });

                if (index == 11) //обрабатывается часть таблицы для кабелей с экраном
                {
                    dbContext.ListCableProperties.Add(new ListCableProperties { PropertyId = 4, CableId = cableRec.Entity.Id });
                    dbContext.SaveChanges();
                }

                recordsCount++;
            }
            else
                throw new Exception($"Не удалось распарсить ячейку таблицы!");
        }

        private string BuildTitle(Cable cable, InsulatedBillet billet, string shield)
        {
            string namePart;
            if (billet.PolymerGroupId == 6)
                namePart = $"КЭВ{shield}Внг(А)-LS ";
            else
                namePart = $"КЭРс{shield}";
            _nameBuilder = new StringBuilder(namePart);

            if (cable.CoverPolymerGroupId == 4)
                _nameBuilder.Append("Пнг(А)-FRHF ");
            if (cable.CoverPolymerGroupId == 5)
                _nameBuilder.Append("Унг(D)-FRHF ");
            namePart = CableCalculations.FormatConductorArea((double)billet.Conductor.AreaInSqrMm);

            _nameBuilder.Append($"{cable.ElementsCount}х{namePart}");
            return _nameBuilder.ToString();
        }
    }
}
