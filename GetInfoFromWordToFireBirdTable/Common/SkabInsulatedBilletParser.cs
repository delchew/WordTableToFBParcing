using System;
using System.Collections.Generic;
using GetInfoFromWordToFireBirdTable.CableEntityes;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class SkabInsulatedBilletParser : ICableDataParcer
    {
        public int ParseDataToDatabase()
        {
            var billetTableProvider = new FirebirdDBTableProvider<CableBillet>();
            billetTableProvider.OpenConnection();

            if (!billetTableProvider.TableExists())
            {
                billetTableProvider.CloseConnection();
                throw new Exception("Таблицы соответствующей размеченному классу не существует в текущей базе данных!");
            }

            int recordsCount = 0;
            var conductorsIdList = new List<int> { 21, 2, 4, 6, 8 }; // 0.5, 0.75, 1.0, 1.5, 2.5
            var polymerGroupsId = new List<int> { 6, 4, 3 }; //6 - PVC-LS, 4 - HF, 3 - Rubber
            var cableShortNamesList = new List<int> { 1, 4 }; //1 - СКАБ250, 4 - СКАБ660

            var cableBillet = new CableBillet();

            foreach (var nameId in cableShortNamesList)
            {
                foreach (var condId in conductorsIdList)
                {
                    foreach (var polymerId in polymerGroupsId)
                    {
                        cableBillet.CableShortNameId = nameId;
                        cableBillet.ConductorId = condId;
                        cableBillet.Diameter = GetDiameter(nameId, condId, polymerId);
                        cableBillet.PolymerGroupId = polymerId;
                        var minThickness = GetMinInsThickness(nameId, polymerId);
                        cableBillet.MinThickness = minThickness;
                        cableBillet.NominalThickness = minThickness + 0.1m;

                        billetTableProvider.AddItem(cableBillet);
                        recordsCount++;
                    }
                }
            }
            billetTableProvider.CloseConnection();
            return recordsCount;
        }

        private decimal GetMinInsThickness(int nameId, int polymerGroupId)
        {
            if (polymerGroupId == 3 || polymerGroupId == 8)
            {
                return nameId == 1 ? 0.6m : 0.9m;
            }
            else
            {
                return nameId == 1 ? 0.4m : 0.6m;
            }
        }

        private Dictionary<int, decimal> PVCDiamDict250 = new Dictionary<int, decimal>
        { {21, 2.0m}, {2, 2.2m}, {4, 2.3m}, {6, 2.7m}, {8, 3.1m} };
        private Dictionary<int, decimal> HFDiamDict250 = new Dictionary<int, decimal>
        { {21, 2.0m}, {2, 2.1m}, {4, 2.3m}, {6, 2.7m}, {8, 3.1m} };
        private Dictionary<int, decimal> RubberDiamDict250 = new Dictionary<int, decimal>
        { {21, 2.3m}, {2, 2.5m}, {4, 2.7m}, {6, 3.0m}, {8, 3.4m} };
        private Dictionary<int, decimal> PVCDiamDict660 = new Dictionary<int, decimal>
        { {21, 2.3m}, {2, 2.5m}, {4, 2.7m}, {6, 3.0m}, {8, 3.6m} };
        private Dictionary<int, decimal> HFDiamDict660 = new Dictionary<int, decimal>
        { {21, 2.3m}, {2, 2.5m}, {4, 2.7m}, {6, 3.0m}, {8, 3.5m} };
        private Dictionary<int, decimal> RubberDiamDict660 = new Dictionary<int, decimal>
        { {21, 2.9m}, {2, 3.1m}, {4, 3.3m}, {6, 3.6m}, {8, 4.0m} };

        public event Action<int, int> ParseReport;

        private decimal GetDiameter(int nameId, int condId, int polymerId)
        {
            if (nameId == 1)
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
