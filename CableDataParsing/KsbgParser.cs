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
    public class KsbgParser : CableParser
    {
        private FirebirdDBProvider _provider;
        private FirebirdDBTableProvider<CablePresenter> _cableTableProvider;
        private FirebirdDBTableProvider<ListCableBilletsPresenter> _listCableBilletsProvider;
        private FirebirdDBTableProvider<ListCablePropertiesPresenter> _listCablePropertiesProvider;

        private IEnumerable<PolymerGroup> _coverPolymerGroups;
        private IEnumerable<InsulatedBillet> _billets;
        private KsbgNameBuilder _nameBuilder;
        private int _recordsCount;

        private FireProtectionClass _fireFRLS;
        private FireProtectionClass _fireFRHF;

        public KsbgParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        {
            _provider = new FirebirdDBProvider(connectionString);
            _cableTableProvider = new FirebirdDBTableProvider<CablePresenter>(_provider);
            _listCableBilletsProvider = new FirebirdDBTableProvider<ListCableBilletsPresenter>(_provider);
            _listCablePropertiesProvider = new FirebirdDBTableProvider<ListCablePropertiesPresenter>(_provider);
        }

        public override int ParseDataToDatabase()
        {
            _recordsCount = 0;

            var configuratorTSimpleAndKG = new TableParserConfigurator(3, 2, 9, 6, 2, 1);
            var configuratorTSimpleAndKGS = new TableParserConfigurator(9, 2, 9, 6, 2, 1);
            var configuratorK = new TableParserConfigurator(3, 2, 7, 6, 2, 1);
            var configuratorKS = new TableParserConfigurator(9, 2, 7, 6, 2, 1);

            var techCond = _dbContext.TechnicalConditions.Where(c => c.Title == "ТУ 16.К99-037-2009").First();

            var voltage = _dbContext.OperatingVoltages.Find(5);

            var climaticModUHL = _dbContext.ClimaticMods.Where(c => c.Title == "УХЛ").First();
            var climaticModV = _dbContext.ClimaticMods.Where(c => c.Title == "В").First();

            var colorOrange = _dbContext.Colors.Where(c => c.Title == "orange").First();
            var colorBlack = _dbContext.Colors.Where(c => c.Title == "black").First();

            _fireFRLS = _dbContext.FireProtectionClasses.Find(18);
            _fireFRHF = _dbContext.FireProtectionClasses.Find(23);

            var polymerLS = _dbContext.PolymerGroups.Find(6);
            var polymerHF = _dbContext.PolymerGroups.Find(4);

            var pair = _dbContext.TwistedElementTypes.Find(2);

            var cablePropertiesSetList = new List<CablePropertySet?>
            {
                CablePropertySet.HasFoilShield,
                CablePropertySet.HasFoilShield | CablePropertySet.HasMicaWinding,
                CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid,
                CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid | CablePropertySet.HasMicaWinding,
                CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube,
                CablePropertySet.HasFoilShield | CablePropertySet.HasArmourBraid | CablePropertySet.HasArmourTube | CablePropertySet.HasMicaWinding,
            };

            _coverPolymerGroups = new[] { polymerLS, polymerHF };

            _billets = _dbContext.InsulatedBillets.Where(b => b.CableBrandName.BrandName == "КСБ")
                                                     .Include(b => b.Conductor)
                                                     .ToList();

            _nameBuilder = new KsbgNameBuilder();

            _wordTableParser = new XceedWordTableParser();
            _wordTableParser.OpenWordDocument(_mSWordFile);

            var tableDataCommon = new List<TableCellData>();
            var cablePresenter = new CablePresenter
            {
                TechCondId = techCond.Id,
                OperatingVoltageId = voltage.Id,
                ClimaticModId = climaticModUHL.Id,
            };

            _provider.OpenConnection();

            try
            {

            }
            finally
            {
                _wordTableParser.CloseWordApp();
                _provider.CloseConnection();
            }

            return _recordsCount;
        }

    }
}
