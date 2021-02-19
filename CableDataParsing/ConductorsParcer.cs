using System;
using CableDataParsing.TableEntityes;
using FirebirdDatabaseProvider;

namespace CableDataParsing
{
    public class ConductorsParcer : ICableDataParcer
    {
        private string _connectionString;
        private FirebirdDBProvider _dBProvider;
        private FirebirdDBTableProvider<ConductorPresenter> _condTableProvider;

        public ConductorsParcer(string connectionString)
        {
            _connectionString = connectionString;
            _dBProvider = new FirebirdDBProvider(_connectionString);
            _condTableProvider = new FirebirdDBTableProvider<ConductorPresenter>(_dBProvider);
        }

        private ConductorPresenter[] _skabConductors = new ConductorPresenter[]
        {
              new ConductorPresenter { Title = "7х0.20 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.3m, ConductorDiameter = 0.6m, AreaInSqrMm = 0.2m },
              new ConductorPresenter { Title = "7х0.26 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.26m, ConductorDiameter = 0.78m, AreaInSqrMm = 0.35m },
              new ConductorPresenter { Title = "16х0.20 ММЛ", Class = 2, MetalId = 2, WiresCount = 16, WiresDiameter = 0.2m, ConductorDiameter = 0.94m, AreaInSqrMm = 0.5m },
              new ConductorPresenter { Title = "14х0.26 ММЛ", Class = 2, MetalId = 2, WiresCount = 14, WiresDiameter = 0.26m, ConductorDiameter = 1.15m, AreaInSqrMm = 0.75m },
              new ConductorPresenter { Title = "19х0.26 ММЛ", Class = 2, MetalId = 2, WiresCount = 19, WiresDiameter = 0.26m, ConductorDiameter = 1.3m, AreaInSqrMm = 1m },
              new ConductorPresenter { Title = "12х0.40 ММЛ", Class = 2, MetalId = 2, WiresCount = 12, WiresDiameter = 0.4m, ConductorDiameter = 1.66m, AreaInSqrMm = 1.5m },
              new ConductorPresenter { Title = "19х0.40 ММЛ", Class = 2, MetalId = 2, WiresCount = 19, WiresDiameter = 0.4m, ConductorDiameter = 2m, AreaInSqrMm = 2.5m }
        };
        
        public event Action<int, int> ParseReport;

        public int ParseDataToDatabase()
        {
            _dBProvider.OpenConnection();

            int recordsCount = 0;
            foreach(var conductor in _skabConductors)
            {
                _condTableProvider.AddItem(conductor);
                recordsCount++;
                ParseReport?.Invoke(_skabConductors.Length, recordsCount);
            }

            _dBProvider.CloseConnection();
            return recordsCount;
        }
    }
}
