using System;
using System.Collections.Generic;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing
{
    public class Kevv_KerspBilletParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;
        public int ParseDataToDatabase()
        {
            int recordsCount = 0;
            var conductorsIdList = new List<int> { 1, 3, 5, 7, 9 }; // заполнить правильно!
            var cableShortnames = new List<int> { 3, 4 }; //3 - КЭВЭВ, 4 - КЭРсЭ

            var cableBillet = new InsulatedBillet
            {
                OperatingVoltageId = 4
            };
            decimal minThickness;
            using (var dbContext = new CablesContext())
            {
                foreach (var shortNameId in cableShortnames)
                {
                    cableBillet.CableShortNameId = shortNameId;
                    cableBillet.PolymerGroupId = shortNameId == 3 ? 6 : 3; //Для КЭВЭВ - пвх LS, для КЭРсЭ - резина

                }
                foreach (var condId in conductorsIdList)
                {
                    cableBillet.ConductorId = condId;
                    cableBillet.Diameter = GetDiameter(voltageId, condId, polymerId);
                    minThickness = GetMinInsThickness(voltageId, polymerId);
                    cableBillet.MinThickness = minThickness;
                    cableBillet.NominalThickness = minThickness + 0.1m;

                    dbContext.InsulatedBillets.Add(cableBillet);

                    recordsCount++;
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
