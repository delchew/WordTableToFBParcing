using CableDataParsing.TableEntityes;
using Cables.Common;
using System.Text;

namespace CableDataParsing.NameBuilders
{
    public class KpseNameBuilder
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public string GetCableName(CablePresenter cable, decimal areaInSqrMm, CablePropertySet? cableProperty)
        {
            _nameBuilder.Clear();
            _nameBuilder.Append("КПС");

            if (cableProperty.HasValue)
            {
                _nameBuilder.Append((cableProperty.Value & CablePropertySet.HasFoilShield) == CablePropertySet.HasFoilShield ? "Э" : string.Empty);
                _nameBuilder.Append((cableProperty.Value & CablePropertySet.HasMicaWinding) == CablePropertySet.HasMicaWinding ? "С" : string.Empty);
            }

            _nameBuilder.Append(cable.CoverPolimerGroupId == 6 ? "нг(А)-FRLS" : "нг(А)-FRHF");

            var namePart = cable.TwistedElementTypeId == 2 ? "2х" : string.Empty;
            var formattedArea = CableCalculations.FormatConductorArea(areaInSqrMm);
            _nameBuilder.Append($" {cable.ElementsCount}х{namePart}{formattedArea}");

            return _nameBuilder.ToString();
        }
    }
}
