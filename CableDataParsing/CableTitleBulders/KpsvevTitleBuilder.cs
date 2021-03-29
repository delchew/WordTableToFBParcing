using CablesDatabaseEFCoreFirebird.Entities;
using System.Linq;
using System.Text;

namespace CableDataParsing.CableTitleBulders
{
    public class KpsvevTitleBuilder : ICableTitleBuilder
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public string GetCableTitle(Cable cable, Cables.Common.CableProperty? cableProperty)
        {
            _nameBuilder.Clear();
            _nameBuilder.Append("КПСВ");
            if (cableProperty != null && (cableProperty & Cables.Common.CableProperty.HasFoilShield) == Cables.Common.CableProperty.HasFoilShield)
                _nameBuilder.Append("Э");
            string namePart = string.Empty;
            switch (cable.CoverPolymerGroup.Title)
            {
                case "PVC":
                case "PVC Term":
                case "PVC LS":
                case "PVC Cold":
                    namePart = "В";
                    break;
                case "PE Self extinguish":
                    namePart = "Пс";
                    break;
            }
            _nameBuilder.Append(namePart);

            if (cableProperty != null && (cableProperty & Cables.Common.CableProperty.HasArmourBraid) == Cables.Common.CableProperty.HasArmourBraid)
            {
                if ((cableProperty & Cables.Common.CableProperty.HasArmourTube) == Cables.Common.CableProperty.HasArmourTube)
                    _nameBuilder.Append($"К{namePart}");
                else
                    _nameBuilder.Append("КГ");
            }
            if (cableProperty != null && (cableProperty & Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube) ==
                (Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube))
            {
                _nameBuilder.Append($"Б{namePart}");
            }

            if (cable.CoverPolymerGroup.Title == "PVC Term")
                _nameBuilder.Append("т");
            if (cable.CoverPolymerGroup.Title == "PVC Cold")
                _nameBuilder.Append("м");
            if (cable.CoverPolymerGroup.Title == "PVC LS")
                _nameBuilder.Append("нг(А)-LS");

            var cableConductorArea = cable.ListCableBillets.First().Billet.Conductor.AreaInSqrMm;
            namePart = Cables.Common.CableCalculations.FormatConductorArea((double)cableConductorArea);
            _nameBuilder.Append($" {cable.ElementsCount}х2х{namePart}");

            return _nameBuilder.ToString();
        }
    }
}
