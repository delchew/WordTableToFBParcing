using CablesDatabaseEFCoreFirebird.Entities;
using Cables.Common;
using System.Text;

namespace CableDataParsing.CableBulders
{
    public class KpsvevTitleBuilder : ICableTitleBuilder
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public string GetCableTitle(Cable cable, InsulatedBillet mainBillet, CablePropertySet? cableProperty, object parameter = null)
        {
            _nameBuilder.Clear();
            if (cable.CoverPolymerGroup.Title == "PVC LSLTx")
            _nameBuilder.Append("ЛОУТОКС ");

                _nameBuilder.Append("КПСВ");
            if (cableProperty != null && (cableProperty & CablePropertySet.HasFoilShield) == CablePropertySet.HasFoilShield)
                _nameBuilder.Append("Э");
            string namePart = string.Empty;
            switch (cable.CoverPolymerGroup.Title)
            {
                case "PVC":
                case "PVC Term":
                case "PVC LS":
                case "PVC Cold":
                case "PVC LSLTx":
                    namePart = "В";
                    break;
                case "PE Self Extinguish":
                    namePart = "Пс";
                    break;
            }
            _nameBuilder.Append(namePart);

            if (cableProperty != null && (cableProperty & CablePropertySet.HasArmourBraid) == CablePropertySet.HasArmourBraid)
            {
                if ((cableProperty & CablePropertySet.HasArmourTube) == CablePropertySet.HasArmourTube)
                    _nameBuilder.Append($"К{namePart}");
                else
                    _nameBuilder.Append("КГ");
            }
            if (cableProperty != null && (cableProperty & CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube) ==
                (CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube))
            {
                _nameBuilder.Append($"Б{namePart}");
            }

            if (cable.CoverPolymerGroup.Title == "PVC Term")
                _nameBuilder.Append("т");
            if (cable.CoverPolymerGroup.Title == "PVC Cold")
                _nameBuilder.Append("м");
            if (cable.CoverPolymerGroup.Title == "PVC LS")
                _nameBuilder.Append("нг(А)-LS");
            if (cable.CoverPolymerGroup.Title == "PVC LSLTx")
                _nameBuilder.Append("нг(А)-LSLTx");

            var cableConductorArea = mainBillet.Conductor.AreaInSqrMm;
            namePart = CableCalculations.FormatConductorArea(cableConductorArea);
            _nameBuilder.Append($" {cable.ElementsCount}х2х{namePart}");

            return _nameBuilder.ToString();
        }
    }
}
