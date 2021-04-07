using System.Collections.Generic;
using System.Linq;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public class BilletParser 
    {
        private string _connectionString;
        private CablesContext _dbContext;

        public BilletParser(string connectionString)
        {
            _connectionString = connectionString;
        }

        public int ParseKevv_Kers()
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

            using (var _dbContext = new CablesContext(_connectionString))
            {
                _dbContext.AddRange(kevvBillets);
                _dbContext.AddRange(kersBillets);
                _dbContext.SaveChanges();
            }
            return recordsCount;
        }

        public void ParseKip()
        {
            var cond078 = _dbContext.Conductors.Where(c => c.WiresDiameter == 0.26m && c.WiresCount == 7 && c.MetalId == 2).Single();
            var cond060 = _dbContext.Conductors.Where(c => c.WiresDiameter == 0.20m && c.WiresCount == 7 && c.MetalId == 2).Single();
            var polymerGroupPE = _dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PE").Single();
            var polymerGroupPVC = _dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC").Single();

            var operatingVoltage = _dbContext.OperatingVoltages.Add(new OperatingVoltage { ACVoltage = 300, ACFriquency = 50, Description = "Переменное - 300В, 50Гц" }).Entity;
            var cableShortName = _dbContext.CableShortNames.Add(new CableShortName { ShortName = "КИП" }).Entity;
            var foamPE = _dbContext.PolymerGroups.Add(new PolymerGroup { Title = "Foamed PE", TitleRus = "Вспененный полиэтилен" }).Entity;

            var kipBillets = new List<InsulatedBillet>
                {
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond060,
                        Diameter = 1.42m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPE
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond060,
                        Diameter = 1.50m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPE
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond078,
                        Diameter = 1.78m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = foamPE
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = cond078,
                        Diameter = 1.60m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPVC
                    }
                };

            _dbContext.InsulatedBillets.AddRange(kipBillets);
            _dbContext.SaveChanges();
        }

        public int ParseKpsvev()
        {
            var conuctor05 = _dbContext.Conductors.Find(17);
            var conuctor075 = _dbContext.Conductors.Find(18);
            var conuctor1 = _dbContext.Conductors.Find(19);
            var conuctor15 = _dbContext.Conductors.Find(20);
            var conuctor25 = _dbContext.Conductors.Find(21);

            var cableShortName = _dbContext.CableShortNames.Find(6); // КПСВ(Э)
            var operatingVoltage = _dbContext.OperatingVoltages.Find(6); // 300В 50Гц, постоянка - до 500В

            var polymerGroups = new List<PolymerGroup>
            {
                _dbContext.PolymerGroups.Find(1), //PVC
                _dbContext.PolymerGroups.Find(6), //PVC LS
                _dbContext.PolymerGroups.Find(7), //PVC LSLTx
                _dbContext.PolymerGroups.Find(11) //PVC term
            };

            List<InsulatedBillet> kpsvvBillets;

            foreach (var group in polymerGroups)
            {
                kpsvvBillets = new List<InsulatedBillet>
                {
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor05,
                        Diameter = 1.8m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor075,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.06m : 1.96m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor1,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.29m : 2.19m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor15,
                        Diameter = 2.66m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor25,
                        Diameter = 3.06m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    }
                };
                _dbContext.AddRange(kpsvvBillets);
            }
            return _dbContext.SaveChanges();
        }

        public int ParseSkab()
        {
            int recordsCount = 0;
            var conductorsIdList = new List<int> { 1, 3, 5, 7, 9 }; // 0.5, 0.75, 1.0, 1.5, 2.5
            var polymerGroupsId = new List<int> { 6, 4, 3 }; //6 - PVC-LS, 4 - HF, 3 - Rubber
            var skabOperatingVoltageIdList = new List<int> { 2, 3 }; //2 - СКАБ250, 3 - СКАБ660

            var cableBillet = new InsulatedBillet();
            decimal minThickness;
            using (var _dbContext = new CablesContext(_connectionString))
            {
                foreach (var voltageId in skabOperatingVoltageIdList)
                {
                    foreach (var condId in conductorsIdList)
                    {
                        foreach (var polymerId in polymerGroupsId)
                        {
                            cableBillet.CableShortNameId = 2;
                            cableBillet.OperatingVoltageId = voltageId;
                            cableBillet.ConductorId = condId;
                            cableBillet.Diameter = GetSkabDiameter(voltageId, condId, polymerId);
                            cableBillet.PolymerGroupId = polymerId;
                            minThickness = GetSkabMinInsThickness(voltageId, polymerId);
                            cableBillet.MinThickness = minThickness;
                            cableBillet.NominalThickness = minThickness + 0.1m;

                            _dbContext.InsulatedBillets.Add(cableBillet);
                            recordsCount++;
                        }
                    }
                }
                _dbContext.SaveChanges();
            }
            return recordsCount;
        }

        private decimal GetSkabMinInsThickness(int voltageId, int polymerGroupId)
        {
            if (polymerGroupId == 3 || polymerGroupId == 8)
            {
                return voltageId == 2 ? 0.6m : 0.9m;
            }
            else
            {
                return voltageId == 2 ? 0.4m : 0.6m;
            }
        }

        private Dictionary<int, decimal> PVCDiamDict250 = new Dictionary<int, decimal>
        { {1, 2.0m}, {3, 2.2m}, {5, 2.3m}, {7, 2.7m}, {9, 3.1m} };
        private Dictionary<int, decimal> HFDiamDict250 = new Dictionary<int, decimal>
        { {1, 2.0m}, {3, 2.1m}, {5, 2.3m}, {7, 2.7m}, {9, 3.1m} };
        private Dictionary<int, decimal> RubberDiamDict250 = new Dictionary<int, decimal>
        { {1, 2.3m}, {3, 2.5m}, {5, 2.7m}, {7, 3.0m}, {9, 3.4m} };
        private Dictionary<int, decimal> PVCDiamDict660 = new Dictionary<int, decimal>
        { {1, 2.3m}, {3, 2.5m}, {5, 2.7m}, {7, 3.0m}, {9, 3.6m} };
        private Dictionary<int, decimal> HFDiamDict660 = new Dictionary<int, decimal>
        { {1, 2.3m}, {3, 2.5m}, {5, 2.7m}, {7, 3.0m}, {9, 3.5m} };
        private Dictionary<int, decimal> RubberDiamDict660 = new Dictionary<int, decimal>
        { {1, 2.9m}, {3, 3.1m}, {5, 3.3m}, {7, 3.6m}, {9, 4.0m} };

        private decimal GetSkabDiameter(int voltageId, int condId, int polymerId)
        {
            if (voltageId == 2)
            {
                switch (polymerId) 
                {
                    case 4: return HFDiamDict250[condId];
                    case 3: return RubberDiamDict250[condId];
                    default: return PVCDiamDict250[condId];
                };
            }
            else
            {
                switch (polymerId)
                {
                    case 4: return HFDiamDict660[condId];
                    case 3: return RubberDiamDict660[condId];
                    default: return PVCDiamDict660[condId];
                }
            }
        }
    }
}
