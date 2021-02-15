using System;
using System.Collections.Generic;
using FirebirdDatabaseProvider;
using CableDataParsing.TableEntityes;

namespace CableDataParsing
{
    public class SkabInsulatedBilletParser : ICableDataParcer
    {
        private string _connectionString;
        private FirebirdDBProvider _dBProvider;
        private FirebirdDBTableProvider<CableBilletPresenter> _billetTableProvider;

        public SkabInsulatedBilletParser(string connectionString)
        {
            _connectionString = connectionString;
            _dBProvider = new FirebirdDBProvider(_connectionString);
            _billetTableProvider = new FirebirdDBTableProvider<CableBilletPresenter>(_dBProvider);
        }

        public int ParseDataToDatabase()
        {
            _dBProvider.OpenConnection();

            if (!_billetTableProvider.TableExists())
            {
                _dBProvider.CloseConnection();
                throw new Exception("Таблицы соответствующей размеченному классу не существует в текущей базе данных!");
            }

            int recordsCount = 0;
            var conductorsIdList = new List<int> { 1, 3, 5, 7, 9 }; // 0.5, 0.75, 1.0, 1.5, 2.5
            var polymerGroupsId = new List<int> { 6, 4, 3 }; //6 - PVC-LS, 4 - HF, 3 - Rubber
            var skabOperatingVoltageIdList = new List<int> { 2, 3 }; //2 - СКАБ250, 3 - СКАБ660

            var cableBillet = new CableBilletPresenter();
            decimal minThickness;
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

                        _billetTableProvider.AddItem(cableBillet);
                        recordsCount++;
                    }
                }
            }
            _dBProvider.CloseConnection();
            return recordsCount;
        }

        private decimal GetMinInsThickness(int voltageId, int polymerGroupId)
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

        public event Action<int, int> ParseReport;

        private decimal GetDiameter(int voltageId, int condId, int polymerId)
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
