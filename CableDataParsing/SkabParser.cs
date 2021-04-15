using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using CableDataParsing.MSWordTableParsers;
using Cables.Common;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;
using FirebirdDatabaseProvider;
using CableDataParsing.TableEntityes;
using CableDataParsing.NameBuilders;
using System.Globalization;

namespace CableDataParsing
{
    public class SkabParser : CableParser
    {
        private FirebirdDBProvider _provider;
        private FirebirdDBTableProvider<CablePresenter> _cableTableProvider;
        private FirebirdDBTableProvider<ListCableBilletsPresenter> _listCableBilletsPresenter;
        private FirebirdDBTableProvider<ListCablePropertiesPresenter> _listCablePropertiesPresenter;

        public SkabParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _provider = new FirebirdDBProvider(connectionString);
            _cableTableProvider = new FirebirdDBTableProvider<CablePresenter>(_provider);
            _listCableBilletsPresenter = new FirebirdDBTableProvider<ListCableBilletsPresenter>(_provider);
            _listCablePropertiesPresenter = new FirebirdDBTableProvider<ListCablePropertiesPresenter>(_provider);
        }

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
                                                     .Include(b => b.OperatingVoltage)
                                                     .ToList();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);

            var maxDiamTableCount = _wordTableParser.DocumentTablesCount / 2;

            var cablePresenter = new CablePresenter { TechCondId = techCond.Id };
            var nameBuilder = new SkabNameBuilder();
            _provider.OpenConnection();
            var tableNumber = 0;
            try
            {
                while (tableNumber < maxDiamTableCount)
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
                                        var tableData = _wordTableParser.GetCableCellsCollection(tableNumber, twistTypeParams.configurator);

                                        foreach (var tableCellData in tableData)
                                        {
                                            if (decimal.TryParse(tableCellData.ColumnHeaderData, NumberStyles.Any, _cultureInfo, out decimal elementsCount) &&
                                                decimal.TryParse(tableCellData.CellData, NumberStyles.Any, _cultureInfo, out decimal maxCoverDiameter) &&
                                                decimal.TryParse(tableCellData.RowHeaderData, NumberStyles.Any, _cultureInfo, out decimal conductorAreaInSqrMm))
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

                                                        cablePresenter.ElementsCount = elementsCount;
                                                        cablePresenter.OperatingVoltageId = voltage.Id;
                                                        cablePresenter.TwistedElementTypeId = twistTypeParams.twistMode.Id;
                                                        cablePresenter.MaxCoverDiameter = maxCoverDiameter;
                                                        cablePresenter.FireProtectionId = matParam.fireClass.Id;
                                                        cablePresenter.CoverPolimerGroupId = matParam.coverPolymerGroup.Id;
                                                        cablePresenter.CoverColorId = (exiParam.HasValue && (!armourType.HasValue || (armourType.Value & CablePropertySet.HasArmourTube) != CablePropertySet.HasArmourTube)) ? colorBlue.Id : colorBlack.Id;
                                                        cablePresenter.ClimaticModId = matParam.coverPolymerGroup.Id == 6 ? climaticModUHL.Id : climaticModV.Id;

                                                        cablePresenter.Title = nameBuilder.GetCableName(cablePresenter, conductorAreaInSqrMm, cableProps);
                                                        var cablePresenterId = _cableTableProvider.AddItem(cablePresenter);

                                                        _listCableBilletsPresenter.AddItem(new ListCableBilletsPresenter { CableId = cablePresenterId, BilletId = billet.Id });

                                                        var intProp = 0b_0000000001;

                                                        for (int j = 0; j < cablePropertiesCount; j++)
                                                        {
                                                            var Prop = (CablePropertySet)intProp;

                                                            if ((cableProps & Prop) == Prop)
                                                            {
                                                                var propertyObj = cablePropertiesList.Where(p => p.BitNumber == intProp).First();
                                                                _listCablePropertiesPresenter.AddItem(new ListCablePropertiesPresenter { PropertyId = propertyObj.Id, CableId = cablePresenterId });

                                                            }
                                                            intProp <<= 1;
                                                        }
                                                        recordsCount++;
                                                    }
                                                }
                                            }
                                            else throw new Exception($"Не удалось распарсить ячейку таблицы №{tableNumber}!");
                                        }
                                    }
                                    OnParseReport((double)(tableNumber + 1) / maxDiamTableCount);
                                    tableNumber++;
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                _wordTableParser.CloseWordApp();
                _provider.CloseConnection();
            }
            return recordsCount;
        }
    }
}
