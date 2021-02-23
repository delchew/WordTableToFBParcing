using CableDataParsing.MSWordTableParsers;
using System;
using Cables.Brands;
using System.Collections.Generic;
using System.IO;
using WordObj = Microsoft.Office.Interop.Word;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using System.Linq;
using Microsoft.EntityFrameworkCore;

namespace CableDataParsing
{
    public class Kevv_KerspParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;

        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;
        private string _connectionString;

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
                pvcBillets = dbContext.InsulatedBillets.Include(b => b.Conductor)
                                                       .Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэв"))
                                                       .ToList();
                rubberBillets = dbContext.InsulatedBillets.Include(b => b.Conductor)
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

                    var dataStartRowIndexes = new int[2] { 4, 12 };

                    List<TableCellData> tableData;
                    List<InsulatedBillet> currentBilletsList;

                    using (var dbContext = new CablesContext(_connectionString))
                    {
                        for (int i = 1; i <= tables.Count; i++)
                        {
                            _wordTableParser.DataColumnsCount = i % 2 == 0 ? 7 : 9;
                            currentBilletsList = i < 3 ? pvcBillets : rubberBillets;
                            foreach (var index in dataStartRowIndexes)
                            {
                                _wordTableParser.DataStartRowIndex = index;
                                tableData = _wordTableParser.GetCableCellsCollection(tables[i]);
                                foreach (var tableCellData in tableData)
                                {
                                    if (int.TryParse(tableCellData.ColumnHeaderData, out int elementsCount) &&
                                        decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                                        decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                                    {

                                        var kevvKersp = new Cable
                                        {
                                            ClimaticModId = 3, //Поставить правильный!
                                            CoverColorId = 1, //Поставить правильный!
                                            CoverPolymerGroupId = 1, //Поставить правильный!
                                            ElementsCount = elementsCount,
                                            FireProtectionClassId = 1, //Поставить правильный!
                                            TwistedElementTypeId = 1, //single
                                            OperatingVoltageId = 3, //Поставить правильный!
                                            TechnicalConditionsId = 1, //Поставить правильный!
                                            MaxCoverDiameter = maxCoverDiameter,
                                            BilletId = (from b in currentBilletsList
                                                        where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm
                                                        select b.Id).First(),
                                            Title = "",
                                        };

                                        var a = dbContext.Cables.Add(kevvKersp);
                                        a.Entity.

                                        recordsCount++;
                                    }
                                    else
                                        throw new Exception($"Не удалось распарсить ячейку таблицы №{i}!");
                                }
                                tableData.Clear();
                                dbContext.SaveChanges();
                                ParseReport(12, )
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
    }
}
