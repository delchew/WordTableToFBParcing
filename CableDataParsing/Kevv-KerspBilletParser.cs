using System;
using System.Collections.Generic;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public class Kevv_KerspBilletParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;

        public void ParseConductors()
        {
            var conductors = new List<Conductor>
            {
                new Conductor{ Title = "7х0,20 ММЛ", WiresCount = 7, MetalId = 2, Class = 4, WiresDiameter = 0.20m, ConductorDiameter = 0.6m, AreaInSqrMm = 0.2m },
                new Conductor{ Title = "7х0,26 ММЛ", WiresCount = 7, MetalId = 2, Class = 4, WiresDiameter = 0.26m, ConductorDiameter = 0.78m, AreaInSqrMm = 0.35m },
                new Conductor{ Title = "16х0,20 ММЛ", WiresCount = 16, MetalId = 2, Class = 4, WiresDiameter = 0.20m, ConductorDiameter = 0.94m, AreaInSqrMm = 0.5m },
                new Conductor{ Title = "14х0,26 ММЛ", WiresCount = 14, MetalId = 2, Class = 4, WiresDiameter = 0.26m, ConductorDiameter = 1.15m, AreaInSqrMm = 0.75m },
                new Conductor{ Title = "19х0,26 ММЛ", WiresCount = 19, MetalId = 2, Class = 4, WiresDiameter = 0.26m, ConductorDiameter = 1.30m, AreaInSqrMm = 1.0m },
                new Conductor{ Title = "12х0,40 ММЛ", WiresCount = 12, MetalId = 2, Class = 4, WiresDiameter = 0.40m, ConductorDiameter = 1.66m, AreaInSqrMm = 1.5m },
                new Conductor{ Title = "19х0,40 ММЛ", WiresCount = 19, MetalId = 2, Class = 4, WiresDiameter = 0.40m, ConductorDiameter = 2.0m, AreaInSqrMm = 2.5m },
            };

            using (var dbContext = new CablesContext())
            {
                dbContext.Conductors.AddRange(conductors);
                var count = dbContext.SaveChanges();
            }
        }

        public int ParseDataToDatabase()
        {
            int recordsCount = 0;

            var kevvBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 22, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 1.2m, CableShortNameId = 3 },
                new InsulatedBillet{ConductorId = 23, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 1.5m, CableShortNameId = 3 },
                new InsulatedBillet{ConductorId = 24, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 1.8m, CableShortNameId = 3 },
                new InsulatedBillet{ConductorId = 25, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 2.3m, CableShortNameId = 3 },
                new InsulatedBillet{ConductorId = 26, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 2.5m, CableShortNameId = 3 },
                new InsulatedBillet{ConductorId = 27, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 3.2m, CableShortNameId = 3 },
                new InsulatedBillet{ConductorId = 28, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 3.6m, CableShortNameId = 3 }
            };

            var kersBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 22, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 1.5m, CableShortNameId = 4 },
                new InsulatedBillet{ConductorId = 23, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 1.8m, CableShortNameId = 4 },
                new InsulatedBillet{ConductorId = 24, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 2.1m, CableShortNameId = 4 },
                new InsulatedBillet{ConductorId = 25, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 2.5m, CableShortNameId = 4 },
                new InsulatedBillet{ConductorId = 26, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 2.7m, CableShortNameId = 4 },
                new InsulatedBillet{ConductorId = 27, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 3.3m, CableShortNameId = 4 },
                new InsulatedBillet{ConductorId = 28, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 3.7m, CableShortNameId = 4 }
            };

            using (var dbContext = new CablesContext())
            {
                dbContext.AddRange(kevvBillets);
                dbContext.SaveChanges();

                dbContext.AddRange(kersBillets);
                dbContext.SaveChanges();
            }
            return recordsCount;
        }
    }
}
