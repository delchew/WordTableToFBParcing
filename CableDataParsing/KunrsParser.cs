using System.Collections.Generic;
using System.IO;
using CableDataParsing.MSWordTableParsers;
using Cables.Common;
using CableDataParsing.TableEntityes;
using System.Linq;
using System;

namespace CableDataParsing
{
    public class KunrsParser : CableParser
    {
        public KunrsParser(string connectionString, FileInfo mSWordFile)
            : base(connectionString, mSWordFile, null) { }
        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;
            var configurator = new TableParserConfigurator().SetDataColumnsCount(4)
                                                         .SetDataRowsCount(8)
                                                         .SetColumnHeadersRowIndex(2)
                                                         .SetRowHeadersColumnIndex(1)
                                                         .SetDataStartColumnIndex(2)
                                                         .SetDataStartRowIndex(3);
            _wordTableParser = new XceedWordTableParser();

            var hasFoilShieldDictionary = new Dictionary<int, bool>
            {
                { 0, false }, { 1, true }, { 2, false }, { 3, true }
            };
            var hasArmourDictionary = new Dictionary<int, bool>
                    {
                        { 0, false }, { 1, false }, { 2, true }, { 3, true }
                    };
            var polimerGroupIdDictionary = new Dictionary<int, int>
                    {
                        { 0, 6 }, { 1, 4 }, { 2, 5 }
                    };
            var insBilletKunrsIdDictionary = new Dictionary<decimal, int>
                    {
                        { 0.75m, 1 }, { 1.0m, 2 }, { 1.5m, 3 }, { 2.5m, 4 }, { 4.0m, 5 }, { 6.0m, 6 }, { 10.0m, 7 }, { 16.0m, 8 }
                    };
            var powerColorsDict = new Dictionary<decimal, PowerWiresColorScheme[]>
                    {
                        { 2m, new [] { PowerWiresColorScheme.N } },
                        { 3m, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none} },
                        { 4m, new [] { PowerWiresColorScheme.N, PowerWiresColorScheme.PE } },
                        { 5m, new [] { PowerWiresColorScheme.PEN, PowerWiresColorScheme.none } }
                    };

            var dictCablePropsToId = new Dictionary<CablePropertySet, long>
                    {
                        { CablePropertySet.HasIndividualFoilShields, 1L },
                        { CablePropertySet.HasIndividualBraidShield, 2L },
                        { CablePropertySet.HasFoilShield, 3L },
                        { CablePropertySet.HasBraidShield, 4L },
                        { CablePropertySet.HasFilling, 5L },
                        { CablePropertySet.HasArmourBraid, 6L },
                        { CablePropertySet.HasArmourTape, 7L },
                        { CablePropertySet.HasArmourTube, 8L },
                        { CablePropertySet.HasWaterBlockStripe, 9L },
                        { CablePropertySet.SparkSafety, 10L }
                    };

            //var kunrs = new KunrsPresenter
            //{
            //    TwistedElementTypeId = 1,
            //    TechCondId = 25,
            //    HasFilling = true,
            //    OperatingVoltageId = 1,
            //    ClimaticModId = 3
            //};

            PowerWiresColorScheme[] powerColorSchemeArray;

            _wordTableParser.OpenWordDocument(_mSWordFile);

            for (int i = 0; i < hasFoilShieldDictionary.Count; i++)
            {
                var tableData = _wordTableParser.GetCableCellsCollection(i, configurator);
                foreach (var tableCellData in tableData)
                {
                    if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                        decimal.TryParse(tableCellData.CellData, out decimal maxCoverDiameter) &&
                        decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
                    {
                        for (int j = 0; j < polimerGroupIdDictionary.Count; j++)
                        {
                            powerColorSchemeArray = powerColorsDict[elementsCount];
                            for (int k = 0; k < powerColorSchemeArray.Length; k++)
                            {
                                throw new NotImplementedException();
                                //recordsCount++;
                            }
                        }
                    }
                }
                OnParseReport((double)(i + 1) / hasFoilShieldDictionary.Count);
                configurator.DataStartRowIndex += configurator.DataRowsCount;
            }
            _wordTableParser.CloseWordApp();
            return recordsCount;
        }
    }
}
