using Cables.Common;
using CablesDatabaseEFCoreFirebird.Entities;
using System.Text;

namespace CableDataParsing.CableBulders
{
    public class Kevv_KerspTitleBuilder : ICableTitleBuilder
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public string GetCableTitle(Cable cable, InsulatedBillet mainBillet, Cables.Common.CableProperty? cableProperty, object parameter = null)
        {
            string shield = string.Empty;
            if(cableProperty.HasValue)
            {
                shield = "Э";
            }
            string namePart;
            if (mainBillet.PolymerGroup.Title == "PVC LS")
                namePart = $"КЭВ{shield}Внг(А)-LS ";
            else
                namePart = $"КЭРс{shield}";
            _nameBuilder = new StringBuilder(namePart);

            if (cable.CoverPolymerGroup.Title == "HFCompound")
                _nameBuilder.Append("Пнг(А)-FRHF ");
            if (cable.CoverPolymerGroup.Title == "PUR")
                _nameBuilder.Append("Унг(D)-FRHF ");
            namePart = CableCalculations.FormatConductorArea((double)mainBillet.Conductor.AreaInSqrMm);

            _nameBuilder.Append($"{cable.ElementsCount}х{namePart}");
            return _nameBuilder.ToString();
        }
    }
}
