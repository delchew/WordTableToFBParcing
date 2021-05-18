using System;
using System.Collections.Generic;
using System.Linq;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public class BilletParser : IDisposable
    {
        private CablesContext _dbContext;

        public BilletParser(string connectionString)
        {
            _dbContext = new CablesContext(connectionString);
        }

        public int AddAllBillets()
        {
            var c1 = AddKunrsBillets();
            var c2 = AddKpseBillets();
            var c3 = AddKevv_KersBillets();
            var c4 = AddKipBillets();
            var c5 = AddKpsvevBillets();
            var c6 = AddSkabBillets();
            var c7 = AddKsbBillets();
            return c1 + c2 + c3 + c4 + c5 + c6 + c7;
        }

        public int AddKsbBillets()
        {
            var ksbBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 31, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 1.95m, MinThickness = null, NominalThickness = null, CableBrandNameId = 8 },
                new InsulatedBillet{ConductorId = 17, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.4m, MinThickness = null, NominalThickness = null, CableBrandNameId = 8 },
                new InsulatedBillet{ConductorId = 18, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.8m, MinThickness = null, NominalThickness = null, CableBrandNameId = 8 },
                new InsulatedBillet{ConductorId = 19, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.8m, MinThickness = null, NominalThickness = null, CableBrandNameId = 8 },
                new InsulatedBillet{ConductorId = 20, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 3.2m, MinThickness = null, NominalThickness = null, CableBrandNameId = 8 },
                new InsulatedBillet{ConductorId = 21, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 3.35m, MinThickness = null, NominalThickness = null, CableBrandNameId = 8 }
            };

            _dbContext.AddRange(ksbBillets);

            return _dbContext.SaveChanges();
        }

        public int AddKunrsBillets()
        {
            var kunrsBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 2, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 2.70m, MinThickness = 0.70m, NominalThickness = 0.90m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 4, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 3.15m, MinThickness = 0.70m, NominalThickness = 0.90m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 6, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 3.30m, MinThickness = 0.70m, NominalThickness = 0.90m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 8, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 3.80m, MinThickness = 0.70m, NominalThickness = 0.90m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 10, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 4.55m, MinThickness = 0.80m, NominalThickness = 1.00m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 12, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 5.20m, MinThickness = 0.80m, NominalThickness = 1.00m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 14, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 6.40m, MinThickness = 0.90m, NominalThickness = 1.10m, CableBrandNameId = 1 },
                new InsulatedBillet{ConductorId = 16, PolymerGroupId = 3, OperatingVoltageId = 1, Diameter = 8.00m, MinThickness = 0.90m, NominalThickness = 1.10m, CableBrandNameId = 1 }
            };

            _dbContext.AddRange(kunrsBillets);

            return _dbContext.SaveChanges();
        }

        public int AddKpseBillets()
        {
            var kpseBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 29, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 1.60m, MinThickness = 0.30m, CableBrandNameId = 7 },
                new InsulatedBillet{ConductorId = 31, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 1.90m, MinThickness = 0.30m, CableBrandNameId = 7 },
                new InsulatedBillet{ConductorId = 17, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.05m, MinThickness = 0.30m, CableBrandNameId = 7 },
                new InsulatedBillet{ConductorId = 18, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.25m, MinThickness = 0.30m, CableBrandNameId = 7 },
                new InsulatedBillet{ConductorId = 19, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.48m, MinThickness = 0.30m, CableBrandNameId = 7 },
                new InsulatedBillet{ConductorId = 20, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 2.73m, MinThickness = 0.30m, CableBrandNameId = 7 },
                new InsulatedBillet{ConductorId = 21, PolymerGroupId = 3, OperatingVoltageId = 5, Diameter = 3.28m, MinThickness = 0.30m, CableBrandNameId = 7 }
            };

            _dbContext.AddRange(kpseBillets);
            
            return _dbContext.SaveChanges();
        }

        public int AddKevv_KersBillets()
        {
            var kevvBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 22, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 1.2m, CableBrandNameId = 3 },
                new InsulatedBillet{ConductorId = 23, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 1.5m, CableBrandNameId = 3 },
                new InsulatedBillet{ConductorId = 24, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 1.8m, CableBrandNameId = 3 },
                new InsulatedBillet{ConductorId = 25, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 2.3m, CableBrandNameId = 3 },
                new InsulatedBillet{ConductorId = 26, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 2.5m, CableBrandNameId = 3 },
                new InsulatedBillet{ConductorId = 27, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 3.2m, CableBrandNameId = 3 },
                new InsulatedBillet{ConductorId = 28, PolymerGroupId = 6, OperatingVoltageId = 4, Diameter = 3.6m, CableBrandNameId = 3 }
            };

            var kersBillets = new List<InsulatedBillet>
            {
                new InsulatedBillet{ConductorId = 22, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 1.5m, CableBrandNameId = 4 },
                new InsulatedBillet{ConductorId = 23, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 1.8m, CableBrandNameId = 4 },
                new InsulatedBillet{ConductorId = 24, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 2.1m, CableBrandNameId = 4 },
                new InsulatedBillet{ConductorId = 25, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 2.5m, CableBrandNameId = 4 },
                new InsulatedBillet{ConductorId = 26, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 2.7m, CableBrandNameId = 4 },
                new InsulatedBillet{ConductorId = 27, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 3.3m, CableBrandNameId = 4 },
                new InsulatedBillet{ConductorId = 28, PolymerGroupId = 3, OperatingVoltageId = 4, Diameter = 3.7m, CableBrandNameId = 4 }
            };

            _dbContext.AddRange(kevvBillets);
            _dbContext.AddRange(kersBillets);
            
            return _dbContext.SaveChanges();
        }

        public int AddKipBillets()
        {
            var cond078 = _dbContext.Conductors.Where(c => c.WiresDiameter == 0.26m && c.WiresCount == 7 && c.MetalId == 2).Single();
            var cond060 = _dbContext.Conductors.Where(c => c.WiresDiameter == 0.20m && c.WiresCount == 7 && c.MetalId == 2).Single();
            var polymerGroupPE = _dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PE").Single();
            var polymerGroupPVC = _dbContext.PolymerGroups.Where(p => p.Title.ToUpper() == "PVC").Single();

            var operatingVoltage = _dbContext.OperatingVoltages.Where(v => v.Description == "Переменное - 300В, 50Гц").First();
            var cableBrandName = _dbContext.CableBrandNames.Where(b => b.BrandName == "КИП").First();
            var foamPE = _dbContext.PolymerGroups.Where(g => g.Title == "Foamed PE").First();

            var kipBillets = new List<InsulatedBillet>
                {
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = cond060,
                        Diameter = 1.42m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPE
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = cond060,
                        Diameter = 1.50m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPE
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = cond078,
                        Diameter = 1.78m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = foamPE
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = cond078,
                        Diameter = 1.60m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = polymerGroupPVC
                    }
                };
            _dbContext.AddRange(kipBillets);
            
            return _dbContext.SaveChanges();
        }

        public int AddKpsvevBillets()
        {
            var conuctor05 = _dbContext.Conductors.Find(17);
            var conuctor075 = _dbContext.Conductors.Find(18);
            var conuctor1 = _dbContext.Conductors.Find(19);
            var conuctor15 = _dbContext.Conductors.Find(20);
            var conuctor25 = _dbContext.Conductors.Find(21);

            var cableBrandName = _dbContext.CableBrandNames.Find(6); // КПСВ(Э)
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
                        CableBrandName = cableBrandName,
                        Conductor = conuctor05,
                        Diameter = 1.8m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = conuctor075,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.06m : 1.96m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = conuctor1,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.29m : 2.19m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
                        Conductor = conuctor15,
                        Diameter = 2.66m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableBrandName = cableBrandName,
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

        public int AddSkabBillets()
        {
            var conductorsIdList = new List<int> { 1, 3, 5, 7, 9 }; // 0.5, 0.75, 1.0, 1.5, 2.5
            var polymerGroupsId = new List<int> { 6, 4, 3 }; //6 - PVC-LS, 4 - HF, 3 - Rubber
            var skabOperatingVoltageIdList = new List<int> { 2, 3 }; //2 - СКАБ250, 3 - СКАБ660

            decimal minThickness;
            foreach (var voltageId in skabOperatingVoltageIdList)
            {
                foreach (var condId in conductorsIdList)
                {
                    foreach (var polymerId in polymerGroupsId)
                    {
                        var cableBillet = new InsulatedBillet
                        {
                            CableBrandNameId = 2,
                            OperatingVoltageId = voltageId,
                            ConductorId = condId,
                            Diameter = GetSkabDiameter(voltageId, condId, polymerId),
                            PolymerGroupId = polymerId
                        };
                        minThickness = GetSkabMinInsThickness(voltageId, polymerId);
                        cableBillet.MinThickness = minThickness;
                        cableBillet.NominalThickness = minThickness + 0.1m;

                        _dbContext.InsulatedBillets.Add(cableBillet);
                    }
                }
            }
            
            return _dbContext.SaveChanges();
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

        public void Dispose()
        {
            _dbContext?.Dispose();
        }
    }
}
