using System;
using GetInfoFromWordToFireBirdTable.TableEntityes;
using FirebirdDatabaseProvider;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class SkabConductorsParcer : ICableDataParcer
    {
        private string _connectionString;
        public SkabConductorsParcer(string connectionString)
        {
            _connectionString = connectionString;
        }

        private ConductorPresenter[] _skabConductors = new ConductorPresenter[]
        {
            new ConductorPresenter { ConductorId = 21, Title = "7х0.30 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.3m, ConductorDiameter = 0.9m, AreaInSqrMm = 0.5m },
            new ConductorPresenter { ConductorId = 2, Title = "7х0.37 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.37m, ConductorDiameter = 1.11m, AreaInSqrMm = 0.75m },
            new ConductorPresenter { ConductorId = 4, Title = "7х0.43 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.43m, ConductorDiameter = 1.29m, AreaInSqrMm = 1.0m },
            new ConductorPresenter { ConductorId = 6, Title = "7х0.53 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.53m, ConductorDiameter = 1.59m, AreaInSqrMm = 1.5m },
            new ConductorPresenter { ConductorId = 8, Title = "7х0.67 ММЛ", Class = 2, MetalId = 2, WiresCount = 7, WiresDiameter = 0.67m, ConductorDiameter = 2.01m, AreaInSqrMm = 2.5m },
        };
        
        public event Action<int, int> ParseReport;

        public int ParseDataToDatabase()
        {
            var conductorTableProvider = new FirebirdDBTableProvider<ConductorPresenter>(_connectionString);
            conductorTableProvider.OpenConnection();

            int recordsCount = 0;
            foreach(var conductor in _skabConductors)
            {
                conductorTableProvider.AddItem(conductor);
                recordsCount++;
                ParseReport?.Invoke(_skabConductors.Length, recordsCount);
            }
           
            conductorTableProvider.CloseConnection();
            return recordsCount;
        }
    }
}
