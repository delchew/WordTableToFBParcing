using System;
using System.Collections.Generic;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing.TableEntityes
{
    public class Kevv_KerspConductorsBuilder : ICableDataParcer
    {
        public event Action<int, int> ParseReport;
        public int ParseDataToDatabase()
        {
            int recordsCount = 0;
            var conductorsIdList = new List<int> { 1, 3, 5, 7, 9 }; // заполнить правильно!
            var polymerGroupsId = new List<int> { 6, 4, 3 }; //заполнить правильно!
            var skabOperatingVoltageIdList = new List<int> { 2, 3 }; //заполнить правильно!

            var cableBillet = new InsulatedBillet();
            decimal minThickness;
            using (var dbContext = new CablesContext())
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
                            cableBillet.Diameter = GetDiameter(voltageId, condId, polymerId);
                            cableBillet.PolymerGroupId = polymerId;
                            minThickness = GetMinInsThickness(voltageId, polymerId);
                            cableBillet.MinThickness = minThickness;
                            cableBillet.NominalThickness = minThickness + 0.1m;

                            dbContext.InsulatedBillets.Add(cableBillet);

                            recordsCount++;
                        }
                    }
                }
                dbContext.SaveChanges();
            }
            return recordsCount;
        }

        private decimal GetMinInsThickness(int voltageId, int polymerGroupId)
        {
            throw new NotImplementedException();
        }


        private decimal GetDiameter(int voltageId, int condId, int polymerId)
        {
            throw new NotImplementedException();
        }
    }
}
