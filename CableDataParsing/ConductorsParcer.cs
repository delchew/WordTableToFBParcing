using System;
using System.Collections.Generic;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public class ConductorsParcer
    {
        private string _connectionString;

        public ConductorsParcer(string connectionString)
        {
            _connectionString = connectionString;
        }

        private List<Conductor> _conductors = new List<Conductor>
        {
            new Conductor{ Title = "7х0,20 ММЛ", WiresCount = 7, MetalId = 2, Class = 4, WiresDiameter = 0.20m, ConductorDiameter = 0.6m, AreaInSqrMm = 0.2m },
            new Conductor{ Title = "7х0,26 ММЛ", WiresCount = 7, MetalId = 2, Class = 4, WiresDiameter = 0.26m, ConductorDiameter = 0.78m, AreaInSqrMm = 0.35m },
            new Conductor{ Title = "16х0,20 ММЛ", WiresCount = 16, MetalId = 2, Class = 4, WiresDiameter = 0.20m, ConductorDiameter = 0.94m, AreaInSqrMm = 0.5m },
            new Conductor{ Title = "14х0,26 ММЛ", WiresCount = 14, MetalId = 2, Class = 4, WiresDiameter = 0.26m, ConductorDiameter = 1.15m, AreaInSqrMm = 0.75m },
            new Conductor{ Title = "19х0,26 ММЛ", WiresCount = 19, MetalId = 2, Class = 4, WiresDiameter = 0.26m, ConductorDiameter = 1.30m, AreaInSqrMm = 1.0m },
            new Conductor{ Title = "12х0,40 ММЛ", WiresCount = 12, MetalId = 2, Class = 4, WiresDiameter = 0.40m, ConductorDiameter = 1.66m, AreaInSqrMm = 1.5m },
            new Conductor{ Title = "19х0,40 ММЛ", WiresCount = 19, MetalId = 2, Class = 4, WiresDiameter = 0.40m, ConductorDiameter = 2.0m, AreaInSqrMm = 2.5m },
        };

        public event Action<int, int> ParseReport;

        public int ParseDataToDatabase()
        {
            int recordsCount = 0;
            using (var dbContext = new CablesContext(_connectionString))
            {
                dbContext.Conductors.AddRange(_conductors);
                var count = dbContext.SaveChanges();
            }
            return recordsCount;
        }
    }
}
