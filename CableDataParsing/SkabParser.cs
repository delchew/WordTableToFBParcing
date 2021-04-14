using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using CableDataParsing.MSWordTableParsers;
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

            var climaticModUHL = _dbContext.ClimaticMods.Find(3);
            var climaticModV = _dbContext.ClimaticMods.Find(7);

            var colorBlack = _dbContext.Colors.Find(2);
            var colorBlue = _dbContext.Colors.Find(3);

            var skabPropertiesList = new List<CablePropertySet?>
            {
                CablePropertySet.HasFilling | CablePropertySet.HasBraidShield,
                CablePropertySet.HasWaterBlockStripe | CablePropertySet.HasFilling | CablePropertySet.HasBraidShield,
                CablePropertySet.HasBraidShield,
                CablePropertySet.HasWaterBlockStripe | CablePropertySet.HasBraidShield,
                CablePropertySet.HasFilling,
                CablePropertySet.HasWaterBlockStripe | CablePropertySet.HasFilling,
                null,
                CablePropertySet.HasWaterBlockStripe
            };

            var operatingVoltages = new List<OperatingVoltage> { voltage250, voltage660 }; // 1 - СКАБ 250, 2 - СКАБ 660 

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
            var polymerRubber = _dbContext.PolymerGroups.Find(3);
            var polymerPUR = _dbContext.PolymerGroups.Find(5);

            var configurator = new TableParserConfigurator().SetDataRowsCount(5)
                                                .SetDataStartColumnIndex(3)
                                                .SetRowHeadersColumnIndex(2);

            var twistParamsList = new List<(TwistedElementType twistMode, CablePropertySet? hasIndividualFoilSHields, TableParserConfigurator configurator)>
            {
                (single, null, new TableParserConfigurator(3, 2, 13, 5, 2, 1)),
                (pair, null, new TableParserConfigurator(10, 2, 14, 5, 9, 1)),
                (triple, null, new TableParserConfigurator(15, 2, 14, 5, 9, 1)),
                (pair, CablePropertySet.HasIndividualFoilShields, new TableParserConfigurator(21, 2, 13, 5, 20, 1)),
                (triple, CablePropertySet.HasIndividualFoilShields, new TableParserConfigurator(26, 2, 13, 5, 20, 1))
            };
            
            var plasticInsParams = new List<(FireProtectionClass fireClass, PolymerGroup insPolymerGroup, PolymerGroup coverPolymerGroup)>
            {
                (fireLS, polymerLS, polymerLS), (fireHF, polymerHF, polymerHF)
            };

            var rubberInsParams = new List<(FireProtectionClass fireClass, PolymerGroup insPolymerGroup, PolymerGroup coverPolymerGroup)>
            {
                (fireFRLS, polymerRubber, polymerLS), (fireFRHF, polymerRubber, polymerHF), (fireCFRHF, polymerRubber, polymerPUR)
            };

            var exiProperties = new List<CablePropertySet?> { null, CablePropertySet.SparkSafety };

            var billets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName == "СКАБ")
                                                     .Include(b => b.Conductor)
                                                     .Include(b => b.PolymerGroup)
                                                     .ToList();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);

            var maxDiamTableCount = _wordTableParser.DocumentTablesCount / 2;

            var patternCable = new Cable { TechnicalConditions = techCond };

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
                                    var tableData = _wordTableParser.GetCableCellsCollection(i, twistTypeParams.configurator);

                                    foreach (var tableCellData in tableData)
                                    {
                                        if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                                            decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                                            decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                                        {
                                            var materialParams = insType == 0 ? plasticInsParams : rubberInsParams;
                                            foreach (var matParam in materialParams)
                                            {
                                                foreach (var exiParam in exiProperties)
                                                {
                                                    var cableProps = CablePropertySet.HasFoilShield;
                                                    if (twistTypeParams.hasIndividualFoilSHields.HasValue)
                                                        cableProps |= twistTypeParams.hasIndividualFoilSHields.Value;
                                                    if (mod.HasValue)
                                                        cableProps |= mod.Value;
                                                    if (armourType.HasValue)
                                                        cableProps |= armourType.Value;
                                                    if (exiParam.HasValue)
                                                        cableProps |= exiParam.Value;

                                                    var billet = billets.Where(b => b.OperatingVoltage == voltage &&
                                                                                    b.PolymerGroup == matParam.insPolymerGroup &&
                                                                                    b.Conductor.AreaInSqrMm == conductorAreaInSqrMm)
                                                                        .First();

                                                    var cable = patternCable.Clone();
                                                    cable.ElementsCount = elementsCount;
                                                    cable.OperatingVoltage = voltage;
                                                    cable.TwistedElementType = twistTypeParams.twistMode;
                                                    cable.MaxCoverDiameter = maxCoverDiameter;
                                                    cable.FireProtectionClass = matParam.fireClass;
                                                    cable.CoverPolymerGroup = matParam.coverPolymerGroup;
                                                    cable.CoverColor = (exiParam.HasValue && (!armourType.HasValue || (armourType.Value & CablePropertySet.HasArmourTube) != CablePropertySet.HasArmourTube)) ? colorBlue : colorBlack;
                                                    cable.ClimaticMod = matParam.coverPolymerGroup.Id == 6 ? climaticModUHL : climaticModV;

                                                    cable.Title = cableTitleBuilder.GetCableTitle(cable, billet, cableProps);

                                                    var cableRec = _dbContext.Cables.Add(cable).Entity;

                                                    _dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet, Cable = cableRec });

                                                    var listOfCableProperties = GetCableAssociatedPropertiesList(cableRec, cableProps);
                                                    _dbContext.ListCableProperties.AddRange(listOfCableProperties);

                                                    _dbContext.SaveChanges();
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
