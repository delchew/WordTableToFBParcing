using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using CableDataParsing.MSWordTableParsers;
using CableDataParsing.TableEntityes;
using Cables.Common;
using CableDataParsing.CableBulders;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;

namespace CableDataParsing
{
    public class SkabParser : CableParser
    {
        public SkabParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile, new SkabTitleBuilder())
        { }

        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;

            var techCond = _dbContext.TechnicalConditions.Where(c => c.Title == "ТУ 16.К99-061-2013").First();

            var voltage250 = _dbContext.OperatingVoltages.Find(2);
            var voltage660 = _dbContext.OperatingVoltages.Find(3);

            //var skabModifycationsList = new List<(bool HasWaterblockStripe, bool HasFilling, bool HasBraidShield)>
            //        {
            //            (false, true, true),
            //            (true, true, true),
            //            (false, false, true),
            //            (true, false, true),
            //            (false, true, false),
            //            (true, true, false),
            //            (false, false, false),
            //            (true, false, false)
            //        };
            var skabPropertiesList = new List<CablePropertySet?>
            {
                CablePropertySet.HasFilling | CablePropertySet.HasBraidShield,
                CablePropertySet.HasWaterBlockStripe | CablePropertySet.HasFilling | CablePropertySet.HasBraidShield,
                CablePropertySet.HasBraidShield | CablePropertySet.HasFoilShield,
                CablePropertySet.HasWaterBlockStripe | CablePropertySet.HasBraidShield,
                CablePropertySet.HasFilling,
                CablePropertySet.HasWaterBlockStripe | CablePropertySet.HasFilling,
                null,
                CablePropertySet.HasWaterBlockStripe
            };

            var operatingVoltages = new List<OperatingVoltage> { voltage250, voltage660 }; // 1 - СКАБ 250, 2 - СКАБ 660 

            //var hasArmourList = new List<(bool hasArmourBraid, bool hasArmourTube)>
            //        {
            //            (hasArmourBraid: false, hasArmourTube: false),
            //            (hasArmourBraid: true, hasArmourTube: false),
            //            (hasArmourBraid: true, hasArmourTube: true)
            //        };
            var skabArmorPropertiesList = new List<CablePropertySet?>
            {
                null,
                CablePropertySet.HasArmourBraid,
                CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube
            };

            var insulationTypes = new List<int> { 0, 1 }; //1 = rubber, 0 = others

            var single = _dbContext.TwistedElementTypes.Find(1);
            var pair = _dbContext.TwistedElementTypes.Find(2);
            var triple = _dbContext.TwistedElementTypes.Find(3);

            var fireLS = _dbContext.FireProtectionClasses.Find(8);
            var fireHF = _dbContext.FireProtectionClasses.Find(13);
            var fireFRLS = _dbContext.FireProtectionClasses.Find(18);
            var fireFRHF = _dbContext.FireProtectionClasses.Find(23);
            var fireCFRHF = _dbContext.FireProtectionClasses.Find(25);

            var polymerLS = _dbContext.PolymerGroups.Find(6);
            var polymerHF = _dbContext.PolymerGroups.Find(4);
            var polymerRupper = _dbContext.PolymerGroups.Find(3);
            var polymerPUR = _dbContext.PolymerGroups.Find(5);

            var configurator = new TableParserConfigurator().SetDataRowsCount(5)
                                                .SetDataStartColumnIndex(3)
                                                .SetRowHeadersColumnIndex(2);


            //var twistTypesParamsList = new List<(TwistedElementType twistMode, CablePropertySet? hasIndividualFoilSHields, int dataStartRowIndex, int dataColumnsCount, int ColumnHeadersRowIndex)>
            //        {
            //            (single, null, 4, 13, 3),
            //            (pair, null, 11, 14, 10),
            //            (triple, null, 16, 14, 10),
            //            (pair, CablePropertySet.HasIndividualFoilShields, 22, 13, 21),
            //            (triple, CablePropertySet.HasIndividualFoilShields, 27, 13, 21)
            //        };

            var twistParamsList = new List<(TwistedElementType twistMode, CablePropertySet? hasIndividualFoilSHields, TableParserConfigurator configurator)>
            {
                (single, null, new TableParserConfigurator(4, 3, 13, 5, 3, 2)),
                (pair, null, new TableParserConfigurator(11, 3, 14, 5, 10, 2)),
                (triple, null, new TableParserConfigurator(16, 3, 14, 5, 10, 2)),
                (pair, CablePropertySet.HasIndividualFoilShields, new TableParserConfigurator(22, 3, 13, 5, 21, 2)),
                (triple, CablePropertySet.HasIndividualFoilShields, new TableParserConfigurator(27, 3, 13, 5, 21, 2))
            };

            var plasticInsMaterialParams = new List<(int fireProtectID, int insPolymerGroupId, int coverPolymerGroupId)>
                    {
                        (8, 6, 6), (13, 4, 4)
                    };
            var rubberInsMaterialParams = new List<(int fireProtectID, int insPolymerGroupId, int coverPolymerGroupId)>
                    {
                        (18, 3, 6), (23, 3, 4), (25, 3, 5)
                    };

            //var exiParams = new List<bool> { false, true };

            var exiProperties = new List<CablePropertySet?> { null, CablePropertySet.SparkSafety };

            var billets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName == "СКАБ")
                                                     .Include(b => b.Conductor)
                                                     .Include(b => b.PolymerGroup)
                                                     .ToList();

            var skabBoolPropertyesList = new List<(bool hasProp, CablePropertySet propType)>();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);
            var maxDiamTableCount = _wordTableParser.DocumentTablesCount / 2;

            var patternCable = new Cable { TechnicalConditionsId = techCond.Id };

            var cableProps = CablePropertySet.HasFoilShield;

            for (int i = 0; i < maxDiamTableCount; i++)
            {
                foreach (var mod in skabPropertiesList)
                {
                    foreach (var voltage in operatingVoltages)
                    {
                        foreach (var insType in insulationTypes)
                        {
                            foreach (var armourType in skabArmorPropertiesList)
                            {
                                foreach (var twistTypeParams in twistParamsList)
                                {
                                    if (twistTypeParams.hasIndividualFoilSHields.HasValue)
                                        cableProps |= twistTypeParams.hasIndividualFoilSHields.Value;
                                    if (mod.HasValue)
                                        cableProps |= mod.Value;
                                    if (armourType.HasValue)
                                        cableProps |= armourType.Value;

                                    var tableData = _wordTableParser.GetCableCellsCollection(i, twistTypeParams.configurator);

                                    foreach (var tableCellData in tableData)
                                    {
                                        if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                                            decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                                            decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                                        {
                                            var materialParams = insType == 0 ? plasticInsMaterialParams : rubberInsMaterialParams;
                                            foreach (var matParam in materialParams)
                                            {
                                                foreach (var exiParam in exiProperties)
                                                {
                                                    if (exiParam.HasValue)
                                                        cableProps |= exiParam.Value;

                                                    var billet = billets.Where(b => b.OperatingVoltage == voltage &&
                                                                                    b.PolymerGroupId == matParam.insPolymerGroupId &&
                                                                                    b.Conductor.AreaInSqrMm == conductorAreaInSqrMm)
                                                                        .First();

                                                    var cable = patternCable.Clone();
                                                    cable.ElementsCount = elementsCount;
                                                    cable.OperatingVoltageId = voltage.Id;
                                                    cable.TwistedElementTypeId = twistTypeParams.twistMode.Id;
                                                    cable.MaxCoverDiameter = maxCoverDiameter;
                                                    cable.FireProtectionClassId = matParam.fireProtectID;
                                                    cable.CoverPolymerGroupId = matParam.coverPolymerGroupId;
                                                    cable.CoverColorId = (exiParam && armourType.hasArmourTube == false) ? 3 : 2;
                                                    cable.ClimaticModId = matParam.coverPolymerGroupId == 6 ? 3 : 7;  //3 - УХЛ, 7 - В

                                                    cable.Title = cableTitleBuilder.GetCableTitle(cable, billet, cableProps);

                                                    ParseTableCellData(cable, tableCellData, new List<InsulatedBillet> { billet }, cableProps);
                                                    recordsCount++;
                                                }
                                            }
                                        }
                                        else throw new Exception($"Не удалось распарсить ячейку таблицы №{i}!");
                                    }
                                }
                                OnParseReport((double)(i + 1) / maxDiamTableCount);
                            }
                        }
                    }
                }
            }
            _wordTableParser.CloseWordApp();
            return recordsCount;
        }
    }
}
