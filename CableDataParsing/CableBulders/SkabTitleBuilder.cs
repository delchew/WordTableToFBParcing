using CablesDatabaseEFCoreFirebird.Entities;
using Cables.Common;
using System.Text;

namespace CableDataParsing.CableBulders
{
    public class SkabTitleBuilder : ICableTitleBuilder
    {
        private readonly StringBuilder _nameBuilder = new StringBuilder();

        public string GetCableTitle(Cable cable, InsulatedBillet mainBillet, CablePropertySet? cableProperty, object parameter = null)
        {
            _nameBuilder.Clear();
            _nameBuilder.Append($"СКАБ ");
            if (cable.OperatingVoltage.ACVoltage == 380 &&
                cable.OperatingVoltage.ACFriquency == 400 &&
                cable.OperatingVoltage.DCVoltage == 540)
                _nameBuilder.Append(250);

            if (cable.OperatingVoltage.ACVoltage == 660 &&
                cable.OperatingVoltage.ACFriquency == 400 &&
                cable.OperatingVoltage.DCVoltage == 1000)
                _nameBuilder.Append(660);        

            if ((cableProperty & CablePropertySet.HasArmourTube) == CablePropertySet.HasArmourTube)
                _nameBuilder.Append("К");
            else
                if ((cableProperty & CablePropertySet.HasArmourBraid) == CablePropertySet.HasArmourBraid)
                _nameBuilder.Append("КГ");

            if (cable.CoverPolymerGroup.Title == "PUR")
                _nameBuilder.Append("У");

            string namePart;

            var fireClassDesignation = cable.FireProtectionClass.Designation;
            namePart = fireClassDesignation.Contains("HF") ? "-ХЛ" : string.Empty;
                
            _nameBuilder.Append(fireClassDesignation);
            _nameBuilder.Append(namePart);

            if (cable.TwistedElementType.Id == 1)
                _nameBuilder.Append($" {cable.ElementsCount}х");
            else
            {
                namePart = (cableProperty & CablePropertySet.HasIndividualFoilShields) == CablePropertySet.HasIndividualFoilShields ? "э" : string.Empty;
                _nameBuilder.Append($" {cable.ElementsCount}х{cable.TwistedElementType.Id}{namePart}х");
            }

            namePart = CableCalculations.FormatConductorArea(mainBillet.Conductor.AreaInSqrMm);
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
    }
}
