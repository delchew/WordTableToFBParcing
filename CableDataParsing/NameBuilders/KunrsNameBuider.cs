using System.Text;
using System.Collections.Generic;
using Cables.Common;
using CableDataParsing.TableEntityes;

namespace CableDataParsing.NameBuilders
{
    public class KunrsNameBuider : ICableNameBuilder<KunrsPresenter>
    {
        private readonly StringBuilder _nameBuilder;
        private readonly Dictionary<long, string> _polymerNamePartsDict;


        public KunrsNameBuider(StringBuilder nameBuilder)
        {
            if (_nameBuilder != null)
                _nameBuilder = nameBuilder;
            else
                _nameBuilder = new StringBuilder();

            _polymerNamePartsDict = new Dictionary<long, string>
            {
                { 6L, "В" },
                { 4L, "П" },
                { 5L, "У" }
            };
        }

        public KunrsNameBuider() : this(new StringBuilder())
        { }

        public string GetCableName(KunrsPresenter cable, CableBilletPresenter insBillet = null, ConductorPresenter conductor = null, object parameter = null)
        {
            _nameBuilder.Clear();
            _nameBuilder.Append("КУНРС ");
            _nameBuilder.Append(cable.HasFoilShield ? "Э" : string.Empty);
            var namePart = _polymerNamePartsDict[cable.CoverPolimerGroupId];
            _nameBuilder.Append(cable.HasArmourTube ? $"{namePart}K" : string.Empty);
            _nameBuilder.Append(namePart);
            _nameBuilder.Append(cable.CoverPolimerGroupId == 6 ? "нг(А)-FRLS" : "нг(А)-FRHF");
            namePart = CableCalculations.FormatConductorArea((double)conductor.AreaInSqrMm);
            _nameBuilder.Append($" {cable.ElementsCount}х{namePart}");
            namePart = (string)parameter;
            if (!string.IsNullOrEmpty(namePart))
                _nameBuilder.Append($" {namePart}");
            return _nameBuilder.ToString();
        }
    }
}
