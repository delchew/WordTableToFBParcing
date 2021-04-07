using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CableDataParsing.CableBulders;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public abstract class CableParser : IDisposable
    {
        private readonly List<CableProperty> cablePropertiesList;

        protected static int cablePropertiesCount;
        protected WordTableParser<TableCellData> _wordTableParser;
        protected FileInfo _mSWordFile;
        protected CablesContext _dbContext;
        protected ICableTitleBuilder cableTitleBuilder;

        public event Action<double> ParseReport;

        static CableParser()
        {
            cablePropertiesCount = Enum.GetNames(typeof(Cables.Common.CableProperty)).Count();
        }

        public CableParser (string connectionString, FileInfo mSWordFile, ICableTitleBuilder cableTitleBuilder)
        {
            _mSWordFile = mSWordFile;
            _dbContext = new CablesContext(connectionString);
            this.cableTitleBuilder = cableTitleBuilder;
            cablePropertiesList = _dbContext.CableProperties.ToList();
        }

        public void Dispose()
        {
            _dbContext?.Dispose();
        }

        public abstract int ParseDataToDatabase();

        protected void OnParseReport(double completedPersentage)
        {
            ParseReport?.Invoke(completedPersentage);
        }

        protected IEnumerable<ListCableProperties> GetCableAssociatedPropertiesList(Cable cable, Cables.Common.CableProperty cableProps)
        {
            var propList = new List<ListCableProperties>();

            var intProp = 0b_0000000001;
            for (int i = 0; i < cablePropertiesCount; i++)
            {
                var Prop = (Cables.Common.CableProperty)intProp;

                if ((cableProps & Prop) == Prop)
                {
                    var propertyObj = cablePropertiesList.Where(p => p.BitNumber == (int)Prop).Single();
                    propList.Add(new ListCableProperties { Property = propertyObj, Cable = cable });
                }
                intProp <<= 1;
            }

            return propList;
        }

        protected void ParseTableCellData(Cable cable, TableCellData tableCellData, IEnumerable<InsulatedBillet> currentBilletsList,
                                Cables.Common.CableProperty? cableProps = null, char splitter = ' ')
        {
            if (decimal.TryParse(tableCellData.ColumnHeaderData, out decimal elementsCount) &&
                decimal.TryParse(tableCellData.RowHeaderData, out decimal conductorAreaInSqrMm))
            {
                decimal height = 0m;
                decimal width = 0m;
                decimal? maxCoverDiameter;
                if (decimal.TryParse(tableCellData.CellData, out decimal diameterValue))
                    maxCoverDiameter = diameterValue;
                else
                {
                    var cableSizes = tableCellData.CellData.Split(splitter);
                    if (cableSizes.Length < 2) return;
                    if (cableSizes.Length == 2 &&
                        decimal.TryParse(cableSizes[0], out height) &&
                        decimal.TryParse(cableSizes[1], out width))
                    {
                        maxCoverDiameter = null;
                    }
                    else throw new Exception("Wrong format table cell data!");
                }
                var billet = (from b in currentBilletsList
                              where b.Conductor.AreaInSqrMm == conductorAreaInSqrMm
                              select b).First();
                cable.ElementsCount = elementsCount;
                cable.MaxCoverDiameter = maxCoverDiameter;
                cable.Title = cableTitleBuilder.GetCableTitle(cable, billet, cableProps);

                var cableRec = _dbContext.Cables.Add(cable).Entity;

                _dbContext.ListCableBillets.Add(new ListCableBillets { Billet = billet, Cable = cableRec });

                if (cableProps.HasValue)
                {
                    var listOfCableProperties = GetCableAssociatedPropertiesList(cableRec, cableProps.Value);
                    _dbContext.ListCableProperties.AddRange(listOfCableProperties);
                }
                if (!maxCoverDiameter.HasValue)
                {
                    var flatSize = new FlatCableSize { Height = height, Width = width, Cable = cableRec };
                    _dbContext.FlatCableSizes.Add(flatSize);
                }
                _dbContext.SaveChanges();
            }
            else
                throw new Exception($"Не удалось распарсить ячейку таблицы!");
        }
    }
}
