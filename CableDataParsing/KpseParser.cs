using CableDataParsing.MSWordTableParsers;
using CableDataParsing.NameBuilders;
using CableDataParsing.TableEntityes;
using Cables.Common;
using CablesDatabaseEFCoreFirebird.Entities;
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

        private IEnumerable<PolymerGroup> _coverPolymerGroups;
        private IEnumerable<InsulatedBillet> _billets;
        private KpseNameBuilder _nameBuilder;
        private int _recordsCount;

        private FireProtectionClass _fireFRLS;
        private FireProtectionClass _fireFRHF;

        public KpseParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _provider = new FirebirdDBProvider(connectionString);
            _cableTableProvider = new FirebirdDBTableProvider<CablePresenter>(_provider);
            _listCableBilletsProvider = new FirebirdDBTableProvider<ListCableBilletsPresenter>(_provider);
            _listCablePropertiesProvider = new FirebirdDBTableProvider<ListCablePropertiesPresenter>(_provider);
        }
        public override int ParseDataToDatabase()
        {
            _recordsCount = 0;

            var configuratorTop = new TableParserConfigurator(3, 2, 9, 7, 2, 1);
            var configuratorBottom = new TableParserConfigurator(11, 2, 9, 7, 10, 1);
            var configuratorSingleBillets = new TableParserConfigurator(4, 1, 2, 7, 3, 0);

            var techCond = _dbContext.TechnicalConditions.Where(c => c.Title == "ТУ 16.К99-036-2007").First();

            var voltage = _dbContext.OperatingVoltages.Find(5);

            var climaticModUHL = _dbContext.ClimaticMods.Find(3);

            var colorOrange = _dbContext.Colors.Find(8);

            _fireFRLS = _dbContext.FireProtectionClasses.Find(18);
            _fireFRHF = _dbContext.FireProtectionClasses.Find(23);

            var polymerLS = _dbContext.PolymerGroups.Find(6);
            var polymerHF = _dbContext.PolymerGroups.Find(4);


            var single = _dbContext.TwistedElementTypes.Find(1);
            var pair = _dbContext.TwistedElementTypes.Find(2);

            var cablePropertiesSetList = new List<CablePropertySet?>
            {
                CablePropertySet.HasFoilShield,
                null,
                CablePropertySet.HasFoilShield | CablePropertySet.HasMicaWinding,
                CablePropertySet.HasMicaWinding
            };

            _coverPolymerGroups = new[] { polymerLS, polymerHF };

            _billets = _dbContext.InsulatedBillets.Where(b => b.CableBrandName.BrandName == "КПС(Э)")
                                                     .Include(b => b.Conductor)
                                                     .ToList();

            _nameBuilder = new KpseNameBuilder();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);

            var tableDataCommon = new List<TableCellData>();
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

                    cablePresenter.TwistedElementTypeId = pair.Id;
                    InsertCablesFromTableCellData(tableDataCommon, cablePresenter, cableProps);
                    tableDataCommon.Clear();

                    tableDataCommon.AddRange(_wordTableParser.GetCableCellsCollection(4, configuratorSingleBillets));
                    cablePresenter.TwistedElementTypeId = single.Id;
                    InsertCablesFromTableCellData(tableDataCommon, cablePresenter, cableProps);
                    tableDataCommon.Clear();
                    configuratorSingleBillets.DataStartColumnIndex += (tableIndex + 1) % 2 == 0 ? -2 : 4;

                    OnParseReport((double)(tableIndex + 1) / cablePropertiesSetList.Count);
                    tableIndex++;
                }
            }
            finally
            {
                _wordTableParser.CloseWordApp();
                _provider.CloseConnection();

            }

            return _recordsCount;
        }

        private void InsertCablesFromTableCellData(IEnumerable<TableCellData> data, CablePresenter cablePresenter, CablePropertySet? cableProps)
        {
            foreach (var tableCellData in data)
            {
                if (decimal.TryParse(tableCellData.ColumnHeaderData, NumberStyles.Any, _cultureInfo, out decimal elementsCount) &&
                    decimal.TryParse(tableCellData.CellData, NumberStyles.Any, _cultureInfo, out decimal maxCoverDiameter) &&
                    decimal.TryParse(tableCellData.RowHeaderData, NumberStyles.Any, _cultureInfo, out decimal conductorAreaInSqrMm))
                {
                    foreach (var polymerGroup in _coverPolymerGroups)
                    {
                        cablePresenter.ElementsCount = elementsCount;
                        cablePresenter.MaxCoverDiameter = maxCoverDiameter;
                        cablePresenter.FireProtectionId = polymerGroup.Id == 6 ? _fireFRLS.Id : _fireFRHF.Id;
                        cablePresenter.CoverPolimerGroupId = polymerGroup.Id;
                        cablePresenter.Title = _nameBuilder.GetCableName(cablePresenter, conductorAreaInSqrMm, cableProps);
                        var cablePresenterId = _cableTableProvider.AddItem(cablePresenter);

                        var billet = (from b in _billets where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm select b).Single();

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
                        _recordsCount++;
                    }
                }
                else throw new Exception($"Не удалось распарсить ячейку таблицы!");
            }
        }
    }
}
