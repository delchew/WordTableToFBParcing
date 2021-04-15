using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Cables.Common;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public abstract class CableParser : IDisposable
    {
        protected readonly List<CableProperty> cablePropertiesList;

        protected CultureInfo _cultureInfo;

        protected static int cablePropertiesCount;
        protected WordTableParser<TableCellData> _wordTableParser;
        protected FileInfo _mSWordFile;
        protected CablesContext _dbContext;

        public event Action<double> ParseReport;

        static CableParser()
        {
            cablePropertiesCount = Enum.GetNames(typeof(CablePropertySet)).Count();
        }

        public CableParser (string connectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _dbContext = new CablesContext(connectionString);
            cablePropertiesList = _dbContext.CableProperties.ToList();

            _cultureInfo = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            _cultureInfo.NumberFormat.CurrencyDecimalSeparator = ",";
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

        protected IEnumerable<ListCableProperties> GetCableAssociatedPropertiesList(Cable cable, CablePropertySet cableProps)
        {
            var propList = new List<ListCableProperties>();

            var intProp = 0b_0000000001;
            for (int i = 0; i < cablePropertiesCount; i++)
            {
                var Prop = (CablePropertySet)intProp;

                if ((cableProps & Prop) == Prop)
                {
                    var propertyObj = cablePropertiesList.Where(p => p.BitNumber == (int)Prop).First();
                    propList.Add(new ListCableProperties { Property = propertyObj, Cable = cable });
                }
                intProp <<= 1;
            }

            return propList;
        }

        
    }
}
