using System.Collections.Generic;
using System.IO;
using CableDataParsing.MSWordTableParsers;
using Cables.Common;
using CableDataParsing.TableEntityes;
using System.Linq;
using CableDataParsing.NameBuilders;
using System.Text;
using System;

namespace CableDataParsing
{
    public class KunrsParser : CableParser
    {
        private int _recordsCount;

        public KunrsParser(string connectionString, FileInfo mSWordFile)
            : base(connectionString, mSWordFile, null) { }
        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;

            _wordTableParser = new WordTableParser().SetDataColumnsCount(4)
                                                    .SetDataRowsCount(8)
                                                    .SetColumnHeadersRowIndex(3)
                                                    .SetRowHeadersColumnIndex(2)
                                                    .SetDataStartColumnIndex(3)
                                                    .SetDataStartRowIndex(4);

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

            var dictCablePropsToId = new Dictionary<CableProperty, long>
                    {
                        { CableProperty.HasIndividualFoilShields, 1L },
                        { CableProperty.HasIndividualBraidShield, 2L },
                        { CableProperty.HasFoilShield, 3L },
                        { CableProperty.HasBraidShield, 4L },
                        { CableProperty.HasFilling, 5L },
                        { CableProperty.HasArmourBraid, 6L },
                        { CableProperty.HasArmourTape, 7L },
                        { CableProperty.HasArmourTube, 8L },
                        { CableProperty.HasWaterBlockStripe, 9L },
                        { CableProperty.SparkSafety, 10L }
                    };

            var kunrs = new KunrsPresenter
            {
                TwistedElementTypeId = 1,
                TechCondId = 25,
                HasFilling = true,
                OperatingVoltageId = 1,
                ClimaticModId = 3
            };

            PowerWiresColorScheme[] powerColorSchemeArray;

            _wordTableParser.OpenWordDocument(_mSWordFile);

            for (int i = 0; i < hasFoilShieldDictionary.Count; i++)
            {
                var tableData = _wordTableParser.GetCableCellsCollection(i + 1);
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
                                recordsCount++;
                            }
                        }
                    }
                }
                OnParseReport(hasFoilShieldDictionary.Count, i + 1);
                _wordTableParser.DataStartRowIndex += _wordTableParser.DataRowsCount;
                tableData.Clear();
            }
            _wordTableParser.CloseWordApp();
            return recordsCount;
        }
    }
}
