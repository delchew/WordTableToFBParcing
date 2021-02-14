using System.Text;
using CableDataParsing.TableEntityes;
using Cables.Common;
using Cables.Materials;

namespace CableDataParsing.NameBuilders
{
    public class SkabNameBuilder : ICableNameBuilder<SkabPresenter>
    {
        private readonly StringBuilder _nameBuilder;

        public SkabNameBuilder(StringBuilder nameBuilder)
        {
            if (_nameBuilder != null)
                _nameBuilder = nameBuilder;
            else
                _nameBuilder = new StringBuilder();
        }

        public SkabNameBuilder() : this(new StringBuilder())
        { }

        public string GetCableName(SkabPresenter cable, CableBilletPresenter insBillet = null, ConductorPresenter conductor = null, object parameter = null)
        {
            _nameBuilder.Clear();

            var skabVoltageType = cable.OperatingVoltage.ACVoltage == 380 ? 250 : 660;
            _nameBuilder.Append($"СКАБ {skabVoltageType}");

            if (cable.HasArmourTube)
                _nameBuilder.Append("K");
            else
                if (cable.HasArmourBraid)
                    _nameBuilder.Append("KГ");

            if (cable.CoverPolimerGroupId == (long)PolymerGroup.PUR)
                _nameBuilder.Append("У");

            _nameBuilder.Append(cable.FireProtectionClass.Designation);

            if (cable.CoverPolimerGroupId == (long)PolymerGroup.HFCompound ||
                cable.CoverPolimerGroupId == (long)PolymerGroup.PUR)
                _nameBuilder.Append("-ХЛ");
            var namePart = cable.HasIndividualFoilShields ? "э" : string.Empty;
            _nameBuilder.Append($" {cable.ElementsCount}х{(int)cable.TwistedElementTypeId}{namePart}х");
            namePart = CableCalculations.FormatConductorArea((double)conductor.AreaInSqrMm);
            _nameBuilder.Append(namePart + "л");

            var braidMod = !cable.HasBraidShield ? "ф" : string.Empty;
            var fillMod = !cable.HasFilling ? "о" : string.Empty;
            var waterBlockMod = cable.HasWaterBlockStripe ? "в" : string.Empty;
            _nameBuilder.Append($" {braidMod}{fillMod}{waterBlockMod}");

            _nameBuilder.Append(cable.SparkSafety ? " Ex-i" : string.Empty);

            return _nameBuilder.ToString();
        }
    }
}
