using System.Text;
using CableDataParsing.TableEntityes;
using Cables.Common;
using Cables.Materials;

namespace CableDataParsing.NameBuilders
{
    public class SkabNameBuilder
    {
        private readonly StringBuilder _nameBuilder;

        public SkabNameBuilder()
        {
            _nameBuilder = new StringBuilder();
        }

        public string GetCableName(CablePresenter cable, decimal areaInSqrMm, CablePropertySet? cableProperty)
        {
            _nameBuilder.Clear();

            var skabVoltageType = cable.OperatingVoltageId == 2 ? 250 : 660;
            _nameBuilder.Append($"СКАБ {skabVoltageType}");

            if ((cableProperty & CablePropertySet.HasArmourTube) == CablePropertySet.HasArmourTube)
                _nameBuilder.Append("K");
            else
                if ((cableProperty & CablePropertySet.HasArmourBraid) == CablePropertySet.HasArmourBraid)
                _nameBuilder.Append("KГ");

            if (cable.CoverPolimerGroupId == (long)PolymerGroup.PUR)
                _nameBuilder.Append("У");

            _nameBuilder.Append(GetFireProtectDesignationById(cable.FireProtectionId));

            string namePart;

            if (cable.TwistedElementTypeId == 1L)
                _nameBuilder.Append($" {cable.ElementsCount}х");
            else
            {
                namePart = (cableProperty & CablePropertySet.HasIndividualFoilShields) == CablePropertySet.HasIndividualFoilShields ? "э" : string.Empty;
                _nameBuilder.Append($" {cable.ElementsCount}х{(int)cable.TwistedElementTypeId}{namePart}х");
            }

            namePart = CableCalculations.FormatConductorArea(areaInSqrMm);
            _nameBuilder.Append(namePart + "л");

            var onlyFoilShield = (cableProperty & CablePropertySet.HasBraidShield) != CablePropertySet.HasBraidShield;
            var withoutFilling = (cableProperty & CablePropertySet.HasFilling) != CablePropertySet.HasFilling;
            var hasWaterblockStripe = (cableProperty & CablePropertySet.HasWaterBlockStripe) == CablePropertySet.HasWaterBlockStripe;

            if (onlyFoilShield || withoutFilling || hasWaterblockStripe)
            {
                _nameBuilder.Append(" ");
                if (onlyFoilShield) _nameBuilder.Append("ф");
                if (withoutFilling) _nameBuilder.Append("о");
                if (hasWaterblockStripe) _nameBuilder.Append("в");
            }

            if ((cableProperty & CablePropertySet.SparkSafety) == CablePropertySet.SparkSafety)
                _nameBuilder.Append(" Ex-i");
            return _nameBuilder.ToString();
        }

        private string GetFireProtectDesignationById(long id)
        {
            switch(id)
            {
                case 8: return "нг(А)-LS";
                case 13: return "нг(А)-HF-ХЛ";
                case 18: return "нг(А)-FRLS";
                case 23: return "нг(А)-FRHF-ХЛ";
                default: return "нг(С)-FRHF-ХЛ";
            }
        }
    }
}
