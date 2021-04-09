using CablesDatabaseEFCoreFirebird.Entities;
using Cables.Common;
using System.Text;

namespace CableDataParsing.CableBulders
{
    public class KipTitleBuilder : ICableTitleBuilder
    {
        private StringBuilder _nameBuilder = new StringBuilder();
        public string GetCableTitle(Cable cable, InsulatedBillet mainBillet, CablePropertySet? cableProperty, object parameter)
        {
            var flagBG = (bool)parameter;

            _nameBuilder.Clear();
            _nameBuilder.Append("КИП");
            string namePart, condDiam;

            var cableConductorDiameter = mainBillet.Conductor.ConductorDiameter;

            if (cableConductorDiameter == 0.60m && cable.ElementsCount != 1.5m)
            {
                namePart = "Э";
                condDiam = "0,60";
            }
            else
            {
                namePart = "вЭ";
                condDiam = "0,78";
            }
            _nameBuilder.Append(namePart);
            switch (cable.CoverPolymerGroup.Title)
            {
                case "PVC":
                case "PVC Term":
                case "PVC LS":
                case "PVC Cold":
                    namePart = "В";
                    break;
                case "PE":
                    namePart = "П";
                    break;
                default:
                    namePart = string.Empty;
                    break;
            }
            _nameBuilder.Append(namePart);

            if ((cableProperty & CablePropertySet.HasArmourBraid) == CablePropertySet.HasArmourBraid)
            {
                if ((cableProperty & CablePropertySet.HasArmourTube) == CablePropertySet.HasArmourTube)
                    _nameBuilder.Append($"К{namePart}");
                else
                    _nameBuilder.Append("КГ");
            }
            if ((cableProperty & CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube) ==
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
            if (cable.CoverPolymerGroup.Title.ToUpper() == "HFCOMPOUND")
            {
                if (flagBG)
                    _nameBuilder.Append("нг(А)-БГ");
                else
                    _nameBuilder.Append("нг(А)-HF");
            }

            _nameBuilder.Append($" {cable.ElementsCount}х2х{condDiam}");

            return _nameBuilder.ToString();
        }
    }
}
