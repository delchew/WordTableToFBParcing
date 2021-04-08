﻿using System.IO;
using System.Linq;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using CablesDatabaseEFCoreFirebird.Entities;
using CableDataParsing.CableBulders;
using CableDataParsing.MSWordTableParsers;

namespace CableDataParsing
{
    public class Kevv_KerspParser : CableParser
    {
        private const int tablesCount = 4; // количество таблиц в документе
        private const int CABLE_BRANDS_COUNT = 672;  //672 марки в таблицах
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
                                                   .Include(b => b.PolymerGroup)
                                                   .ToList();
            var rubberBillets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName.ToLower().StartsWith("кэрс"))
                                                  .Include(b => b.Conductor)
                                                  .Include(b => b.PolymerGroup)
                                                  .ToList();

            _wordTableParser = new XceedWordTableParser().SetDataRowsCount(7)
                                                    .SetColumnHeadersRowIndex(2)
                                                    .SetRowHeadersColumnIndex(2)
                                                    .SetDataStartColumnIndex(3);

            var kersParams = new List<(FireProtectionClass fireClass, PolymerGroup polymer, Color color)>
            {
                (fireClassHF, polymerHF, colorBlack), (fireClassPURHF, polymerPUR, colorOrange)
            };

            var cableProps = new List<Cables.Common.CableProperty?>
            {
                null, Cables.Common.CableProperty.HasBraidShield
            };

            _wordTableParser.OpenWordDocument(_mSWordFile);

            var patternCable = new Cable
            {
                TwistedElementType = twistedElementType,
                TechnicalConditions = techCond,
                ClimaticMod = climaticMod,
                OperatingVoltage = operatingVoltage
            };

            for (int i = 0; i < tablesCount; i++)
            {
                _wordTableParser.DataColumnsCount = (i + 1) % 2 == 0 ? 7 : 9; //выбираем число столбцов в зависимости от чётности номера таблицы

                foreach (var prop in cableProps)
                {
                    _wordTableParser.DataStartRowIndex = prop.HasValue ? 10 : 3;

                    var tableData = _wordTableParser.GetCableCellsCollection(i);

                    foreach (var tableCellData in tableData)
                    {
                        
                        if (i < 2) //первые 2 таблицы для КЭВВ, остальные - КЭРс
                        {
                            var cable = patternCable.Clone();
                            cable.FireProtectionClass = fireClassPVCLS;
                            cable.CoverPolymerGroup = polymerPVCLS;
                            cable.CoverColor = colorGrey;
                            
                            ParseTableCellData(cable, tableCellData, pvcBillets, prop);
                            _recordsCount++;
                        }
                        else
                        {
                            foreach (var param in kersParams)
                            {
                                var cable = patternCable.Clone();
                                cable.FireProtectionClass = param.fireClass;
                                cable.CoverPolymerGroup = param.polymer;
                                cable.CoverColor = param.color;

                                ParseTableCellData(cable, tableCellData, rubberBillets, prop);
                                _recordsCount++;
                            }
                        }
                    }
                    OnParseReport((double)_recordsCount / CABLE_BRANDS_COUNT);
                }
            }
            _wordTableParser.CloseWordApp();
            return _recordsCount;
        }
    }
}
