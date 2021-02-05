using System;
using System.Collections.Generic;
using GetInfoFromWordToFireBirdTable.CableEntityes;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class SkabInsulatedBilletParser : CableDataParcer
    {
        public override int ParseDataToDatabase()
        {
            var billetTableProvider = new FirebirdDBTableProvider<CableBillet>();
            if (!billetTableProvider.TableExists())
                throw new Exception("Таблицы соответствующей размеченному классу не существует в текущей базе данных!");

            int recordsCount = 0;
            var conductorsIdList = new List<int> { 21, 2, 4, 6, 8 }; // 0.5, 0.75, 1.0, 1.5, 2.5
            var polymerGroupsId = new List<int> { 6, 4, 3 }; //6 - PVC-LS, 7 - PVC-LSLTx, 4 - HF, 3 - Rubber
            var cableShortNamesList = new List<int> { 1, 4 }; //1 - СКАБ250, 4 - СКАБ660

            var cableBillet = new CableBillet();
            billetTableProvider.OpenConnection();

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
                        cableBillet.NominalThickness = minThickness + 0.1;

                        billetTableProvider.AddItem(cableBillet);
                        recordsCount++;
                    }
                }
            }
            billetTableProvider.CloseConnection();
            return recordsCount;
        }

        private double GetMinInsThickness(int nameId, int polymerGroupId)
        {
            if (polymerGroupId == 3 || polymerGroupId == 8)
            {
                return nameId == 1 ? 0.6 : 0.9;
            }
            else
            {
                return nameId == 1 ? 0.4 : 0.6;
            }
        }

        private Dictionary<int, double> PVCDiamDict250 = new Dictionary<int, double>
        { {21, 2.0}, {2, 2.2}, {4, 2.3}, {6, 2.7}, {8, 3.1} };
        private Dictionary<int, double> HFDiamDict250 = new Dictionary<int, double>
        { {21, 2.0}, {2, 2.1}, {4, 2.3}, {6, 2.7}, {8, 3.1} };
        private Dictionary<int, double> RubberDiamDict250 = new Dictionary<int, double>
        { {21, 2.3}, {2, 2.5}, {4, 2.7}, {6, 3.0}, {8, 3.4} };
        private Dictionary<int, double> PVCDiamDict660 = new Dictionary<int, double>
        { {21, 2.3}, {2, 2.5}, {4, 2.7}, {6, 3.0}, {8, 3.6} };
        private Dictionary<int, double> HFDiamDict660 = new Dictionary<int, double>
        { {21, 2.3}, {2, 2.5}, {4, 2.7}, {6, 3.0}, {8, 3.5} };
        private Dictionary<int, double> RubberDiamDict660 = new Dictionary<int, double>
        { {21, 2.9}, {2, 3.1}, {4, 3.3}, {6, 3.6}, {8, 4.0} };

        public override event Action<int, int> ParseReport;

        private double GetDiameter(int nameId, int condId, int polymerId)
        {
            if (nameId == 1)
            {
                switch (polymerId)
                {
                    case 4: return HFDiamDict250[condId];
                    case 3: return RubberDiamDict250[condId];
                    default: return PVCDiamDict250[condId];
                }
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
