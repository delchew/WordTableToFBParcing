using CableDataParsing.MSWordTableParsers;
using CableDataParsing.NameBuilders;
using CableDataParsing.TableEntityes;
using Cables.Common;
using FirebirdDatabaseProvider;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace CableDataParsing
{
    public class KpseParser : CableParser
    {
        private FirebirdDBProvider _provider;
        private FirebirdDBTableProvider<CablePresenter> _cableTableProvider;
        private FirebirdDBTableProvider<ListCableBilletsPresenter> _listCableBilletsProvider;
        private FirebirdDBTableProvider<ListCablePropertiesPresenter> _listCablePropertiesProvider;

        public KpseParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _provider = new FirebirdDBProvider(connectionString);
            _cableTableProvider = new FirebirdDBTableProvider<CablePresenter>(_provider);
            _listCableBilletsProvider = new FirebirdDBTableProvider<ListCableBilletsPresenter>(_provider);
            _listCablePropertiesProvider = new FirebirdDBTableProvider<ListCablePropertiesPresenter>(_provider);
        }
        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;

            var configuratorTop = new TableParserConfigurator(3, 2, 9, 7, 2, 2);
            var configuratorBottom = new TableParserConfigurator(11, 2, 9, 7, 10, 2);
            var configuratorSingleBillets = new TableParserConfigurator(4, 1, 2, 7, 3, 0);

            var techCond = _dbContext.TechnicalConditions.Where(c => c.Title == "ТУ 16.К99-036-2007").First();

            var voltage = _dbContext.OperatingVoltages.Find(5);

            var climaticModUHL = _dbContext.ClimaticMods.Find(3);

            var colorOrange = _dbContext.Colors.Find(8);

            var fireFRLS = _dbContext.FireProtectionClasses.Find(18);
            var fireFRHF = _dbContext.FireProtectionClasses.Find(23);

            var polymerLS = _dbContext.PolymerGroups.Find(6);
            var polymerHF = _dbContext.PolymerGroups.Find(4);

            var coverPolymerGroups = new[] { polymerLS, polymerHF };

            var single = _dbContext.TwistedElementTypes.Find(1);
            var pair = _dbContext.TwistedElementTypes.Find(2);

            var cablePropertiesSetList = new List<CablePropertySet?>
            {
                CablePropertySet.HasFoilShield,
                null,
                CablePropertySet.HasFoilShield | CablePropertySet.HasMicaWinding,
                CablePropertySet.HasMicaWinding
            };

            var billets = _dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName == "КПС(Э)")
                                                     .Include(b => b.Conductor)
                                                     .ToList();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);

            var tableDataCommon = new List<TableCellData>();
            var nameBuilder = new KpseNameBuilder();
            var cablePresenter = new CablePresenter
            {
                TechCondId = techCond.Id,
                OperatingVoltageId = voltage.Id,
                ClimaticModId = climaticModUHL.Id,
                CoverColorId = colorOrange.Id
            };

            _provider.OpenConnection();

            try
            {
                var tableIndex = 0;
                foreach (var cableProps in cablePropertiesSetList)
                {
                    tableDataCommon.AddRange(_wordTableParser.GetCableCellsCollection(tableIndex, configuratorTop));
                    tableDataCommon.AddRange(_wordTableParser.GetCableCellsCollection(tableIndex, configuratorBottom));

                    foreach (var tableCellData in tableDataCommon)
                    {
                        if (decimal.TryParse(tableCellData.ColumnHeaderData, NumberStyles.Any, _cultureInfo, out decimal elementsCount) &&
                            decimal.TryParse(tableCellData.CellData, NumberStyles.Any, _cultureInfo, out decimal maxCoverDiameter) &&
                            decimal.TryParse(tableCellData.RowHeaderData, NumberStyles.Any, _cultureInfo, out decimal conductorAreaInSqrMm))
                        {
                            foreach (var polymerGroup in coverPolymerGroups)
                            {
                                cablePresenter.ElementsCount = elementsCount;
                                cablePresenter.MaxCoverDiameter = maxCoverDiameter;
                                cablePresenter.FireProtectionId = polymerGroup.Id == 6 ? fireFRLS.Id : fireFRHF.Id;
                                cablePresenter.CoverPolimerGroupId = polymerGroup.Id;
                                cablePresenter.TwistedElementTypeId = pair.Id;
                                cablePresenter.Title = nameBuilder.GetCableName(cablePresenter, conductorAreaInSqrMm, cableProps);
                                var cablePresenterId = _cableTableProvider.AddItem(cablePresenter);

                                var billet = (from b in billets where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm select b).Single();

                                _listCableBilletsProvider.AddItem(new ListCableBilletsPresenter { CableId = cablePresenterId, BilletId = billet.Id });

                                var intProp = 0b_0000000001;

                                for (int j = 0; j < cablePropertiesCount; j++)
                                {
                                    var Prop = (CablePropertySet)intProp;

                                    if ((cableProps & Prop) == Prop)
                                    {
                                        var propertyObj = cablePropertiesList.Where(p => p.BitNumber == intProp).First();
                                        _listCablePropertiesProvider.AddItem(new ListCablePropertiesPresenter { PropertyId = propertyObj.Id, CableId = cablePresenterId });
                                    }
                                    intProp <<= 1;
                                }
                                recordsCount++;
                            }
                        }
                        else throw new Exception($"Не удалось распарсить ячейку таблицы №{tableIndex}!");
                    }

                    tableDataCommon.Clear();
                    OnParseReport((double)tableIndex / cablePropertiesSetList.Count);
                    tableIndex++;
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
