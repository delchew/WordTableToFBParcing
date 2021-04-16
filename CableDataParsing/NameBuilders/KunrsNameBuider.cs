using System.Text;
using System.Collections.Generic;
using Cables.Common;
using CableDataParsing.TableEntityes;
using Cables;

namespace CableDataParsing.NameBuilders
{
    public class KunrsNameBuider
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

        public string GetCableName(CablePresenter cable, decimal areaInSqrMm, CablePropertySet? cableProperty, PowerWiresColorScheme? powerWiresColorScheme)
        {
            _nameBuilder.Clear();
            _nameBuilder.Append("КУНРС ");
            var namePart = _polymerNamePartsDict[cable.CoverPolimerGroupId];
            if (cableProperty.HasValue)
            {
                _nameBuilder.Append((cableProperty & CablePropertySet.HasFoilShield) == CablePropertySet.HasFoilShield ? "Э" : string.Empty);
                _nameBuilder.Append((cableProperty & CablePropertySet.HasArmourTube) == CablePropertySet.HasArmourTube ? $"{namePart}K" : string.Empty);
            }
            _nameBuilder.Append(namePart);
            _nameBuilder.Append(cable.CoverPolimerGroupId == 6 ? "нг(А)-FRLS" : "нг(А)-FRHF");
            namePart = CableCalculations.FormatConductorArea(areaInSqrMm);
            _nameBuilder.Append($" {cable.ElementsCount}х{namePart}");
            if(powerWiresColorScheme.HasValue)
                namePart= powerWiresColorScheme.Value.GetDescription();
            if (!string.IsNullOrEmpty(namePart))
                _nameBuilder.Append($" {namePart}");
            return _nameBuilder.ToString();
        }
    }
}
