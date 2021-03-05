using System;
using System.IO;
using System.Linq;
using CableDataParsing.MSWordTableParsers;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public abstract class CableParser : IDisposable
    {
        public event Action<int, int> ParseReport;
        protected WordTableParser _wordTableParser;
        protected FileInfo _mSWordFile;
        protected CablesContext _dbContext;
        protected int cablePropertiesCount = Enum.GetNames(typeof(Cables.Common.CableProperty)).Count();


        public CableParser (string connectionString, FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
            _dbContext = new CablesContext(connectionString);
        }

        public void Dispose()
        {
            _dbContext?.Dispose();
        }

        public abstract string ParceDataToDatabase();

        protected abstract string GetCableTitle(Cable cable, params object[] cableParametres);

        protected void AddCablePropertiesToDBContext(Cable cable, Cables.Common.CableProperty cableProps)
        {
            if (_dbContext == null)
                throw new NullReferenceException("Объект для работы с базой даных не инициализирован!");

            var intProp = 0b_0000000001;
            var cablePropertiesList = _dbContext.CableProperties.ToList();

            for (int i = 0; i < cablePropertiesCount; i++)
            {
                var Prop = (Cables.Common.CableProperty)intProp;

                if ((cableProps & Prop) == Prop)
                {
                    var propertyObj = cablePropertiesList.Where(p => p.BitNumber == (int)Prop).Single();
                    _dbContext.ListCableProperties.Add(new ListCableProperties { Property = propertyObj, Cable = cable });
                }
                intProp <<= 1;
            }
            _dbContext.SaveChanges();
        }
    }
}
