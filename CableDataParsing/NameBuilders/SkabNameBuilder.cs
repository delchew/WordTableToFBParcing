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

            var skabVoltageType = cable.OperatingVoltageId == 2 ? 250 : 660;
            _nameBuilder.Append($"СКАБ {skabVoltageType}");

            if (cable.HasArmourTube)
                _nameBuilder.Append("K");
            else
                if (cable.HasArmourBraid)
                _nameBuilder.Append("KГ");

            if (cable.CoverPolimerGroupId == (long)PolymerGroup.PUR)
                _nameBuilder.Append("У");

            _nameBuilder.Append(GetFireProtectDesignationById(cable.FireProtectionId));

            string namePart;

            if (cable.TwistedElementTypeId == 1L)
                _nameBuilder.Append($" {cable.ElementsCount}х");
            else
            {
                namePart = cable.HasIndividualFoilShields ? "э" : string.Empty;
                _nameBuilder.Append($" {cable.ElementsCount}х{(int)cable.TwistedElementTypeId}{namePart}х");
            }

            namePart = CableCalculations.FormatConductorArea((double)conductor.AreaInSqrMm);
            _nameBuilder.Append(namePart + "л");

            var braidMod = !cable.HasBraidShield ? "ф" : string.Empty;
            var fillMod = !cable.HasFilling ? "о" : string.Empty;
            var waterBlockMod = cable.HasWaterBlockStripe ? "в" : string.Empty;
            _nameBuilder.Append($" {braidMod}{fillMod}{waterBlockMod}");

            _nameBuilder.Append(cable.SparkSafety ? " Ex-i" : string.Empty);

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
