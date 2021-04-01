using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CableDataParsing.CableTitleBulders;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public abstract class CableParser : IDisposable
    {
        private readonly List<CableProperty> cablePropertiesList;

        protected static int cablePropertiesCount;
        protected WordTableParser _wordTableParser;
        protected FileInfo _mSWordFile;
        protected CablesContext _dbContext;
        protected ICableTitleBuilder cableTitleBuilder;

        public event Action<int, int> ParseReport;

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
    }
}
