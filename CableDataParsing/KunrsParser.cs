using System.Collections.Generic;
using System.IO;
using CableDataParsing.MSWordTableParsers;
using Cables.Common;
using CableDataParsing.TableEntityes;
using System.Linq;
using CableDataParsing.NameBuilders;
using FirebirdDatabaseProvider;
using Microsoft.EntityFrameworkCore;
using System.Globalization;

namespace CableDataParsing
{
    public class KunrsParser : CableParser
    {
        private FirebirdDBProvider _provider;
        private FirebirdDBTableProvider<CablePresenter> _cableTableProvider;
        private FirebirdDBTableProvider<ListCableBilletsPresenter> _listCableBilletsProvider;
        private FirebirdDBTableProvider<ListCablePropertiesPresenter> _listCablePropertiesProvider;
        private FirebirdDBTableProvider<ListCablePowerColorPresenter> _listCablePowerColorProvider;
        public KunrsParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _provider = new FirebirdDBProvider(connectionString);
            _cableTableProvider = new FirebirdDBTableProvider<CablePresenter>(_provider);
            _listCableBilletsProvider = new FirebirdDBTableProvider<ListCableBilletsPresenter>(_provider);
            _listCablePropertiesProvider = new FirebirdDBTableProvider<ListCablePropertiesPresenter>(_provider);
            _listCablePowerColorProvider = new FirebirdDBTableProvider<ListCablePowerColorPresenter>(_provider);
        }
        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;
            var nameBuilder = new KunrsNameBuider();
            var configurator = new TableParserConfigurator().SetDataColumnsCount(4)
                                                         .SetDataRowsCount(8)
                                                         .SetColumnHeadersRowIndex(2)
                                                         .SetRowHeadersColumnIndex(1)
                                                         .SetDataStartColumnIndex(2)
                                                         .SetDataStartRowIndex(3);
            _wordTableParser = new XceedWordTableParser();

            var cablePropsList = new List<CablePropertySet?>
            {
                null,
                CablePropertySet.HasFoilShield,
                CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube,
                CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube
            };

            var polymerGroupIdList = new List<int> { 6, 4, 5 };

            var powerColorsDict = new Dictionary<decimal, PowerWiresColorScheme[]>
            {
                { 2m, new [] { PowerWiresColorScheme.N } },
                { 3m, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none} },
                { 4m, new [] { PowerWiresColorScheme.N, PowerWiresColorScheme.PE } },
                { 5m, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none } }
            };

            var billets = _dbContext.InsulatedBillets.Where(b => b.CableBrandName.BrandName == "КУНРС")
                                                     .Include(b => b.Conductor)
                                                     .ToList();

            var kunrs = new CablePresenter
            {
                TwistedElementTypeId = 1,
                TechCondId = 25,
                OperatingVoltageId = 1,
                ClimaticModId = 3
            };

            PowerWiresColorScheme[] powerColorSchemeArray;

            _wordTableParser.OpenWordDocument(_mSWordFile);
            _provider.OpenConnection();
            try
            {
                var parsePartNumber = 1d;
                foreach (var prop in cablePropsList)
                {
                    var tableData = _wordTableParser.GetCableCellsCollection(0, configurator);
                    foreach (var tableCellData in tableData)
                    {
                        if (decimal.TryParse(tableCellData.ColumnHeaderData, NumberStyles.Any, _cultureInfo, out decimal elementsCount) &&
                            decimal.TryParse(tableCellData.CellData, NumberStyles.Any, _cultureInfo, out decimal maxCoverDiameter) &&
                            decimal.TryParse(tableCellData.RowHeaderData, NumberStyles.Any, _cultureInfo, out decimal conductorAreaInSqrMm))
                        {
                            foreach (var polymerGroupId in polymerGroupIdList)
                            {
                                powerColorSchemeArray = powerColorsDict[elementsCount];
                                foreach (var powerColorScheme in powerColorSchemeArray)
                                {
                                    var cableProps = CablePropertySet.HasFilling;
                                    if (prop.HasValue)
                                        cableProps |= prop.Value;

                                    kunrs.ElementsCount = elementsCount;
                                    kunrs.MaxCoverDiameter = maxCoverDiameter;
                                    kunrs.FireProtectionId = polymerGroupId == 6 ? 18 : 23;
                                    kunrs.CoverPolimerGroupId = polymerGroupId;
                                    kunrs.CoverColorId = polymerGroupId == 5 ? 8 : 2;
                                    var billet = (from b in billets
                                                 where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm
                                                 select b).Single();
                                    kunrs.Title = nameBuilder.GetCableName(kunrs, conductorAreaInSqrMm, cableProps, powerColorScheme);
                                    var cableId = _cableTableProvider.AddItem(kunrs);

                                    _listCableBilletsProvider.AddItem(new ListCableBilletsPresenter { CableId = cableId, BilletId = billet.Id });
                                    _listCablePowerColorProvider.AddItem(new ListCablePowerColorPresenter { CableId = cableId, PowerColorSchemeId = (int)powerColorScheme });

                                    var intProp = 0b_0000000001;

                                    for (int m = 0; m < cablePropertiesCount; m++)
                                    {
                                        var Prop = (CablePropertySet)intProp;

                                        if ((cableProps & Prop) == Prop)
                                        {
                                            var propertyObj = cablePropertiesList.Where(p => p.BitNumber == intProp).First();
                                            _listCablePropertiesProvider.AddItem(new ListCablePropertiesPresenter { PropertyId = propertyObj.Id, CableId = cableId });
                                        }
                                        intProp <<= 1;
                                    }
                                    recordsCount++;
                                }
                            }
                        }
                    }
                    OnParseReport(parsePartNumber / cablePropsList.Count);
                    configurator.DataStartRowIndex += configurator.DataRowsCount;
                    parsePartNumber++;
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
