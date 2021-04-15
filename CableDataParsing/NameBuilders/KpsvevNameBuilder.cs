using System.Text;
using CableDataParsing.TableEntityes;
using Cables.Common;

namespace CableDataParsing.NameBuilders
{
    public class KpsvevNameBuilder
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public string GetCableTitle(CablePresenter cable, decimal areaInSqrMm, CablePropertySet? cableProperty, object parameter = null)
        {
            _nameBuilder.Clear();
            if (cable.CoverPolimerGroupId == 7) //PVC LSLTx
                _nameBuilder.Append("ЛОУТОКС ");

            _nameBuilder.Append("КПСВ");
            if (cableProperty != null && (cableProperty & CablePropertySet.HasFoilShield) == CablePropertySet.HasFoilShield)
                _nameBuilder.Append("Э");
            string namePart = string.Empty;
            switch (cable.CoverPolimerGroupId)
            {
                case 1: //PVC
                case 11: //PVC Term
                case 6: //PVC LS
                case 10: //PVC Cold
                case 7: //PVC LSLTx
                    namePart = "В";
                    break;
                case 12: //PE Self Extinguish
                    namePart = "Пс";
                    break;
            }
            _nameBuilder.Append(namePart);

            if (cableProperty != null && (cableProperty & CablePropertySet.HasArmourBraid) == CablePropertySet.HasArmourBraid)
            {
                if ((cableProperty & CablePropertySet.HasArmourTube) == CablePropertySet.HasArmourTube)
                    _nameBuilder.Append($"К{namePart}");
                else
                    _nameBuilder.Append("КГ");
            }
            if (cableProperty != null && (cableProperty & CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube) ==
                (CablePropertySet.HasArmourTape | CablePropertySet.HasArmourTube))
            {
                _nameBuilder.Append($"Б{namePart}");
            }

            if (cable.CoverPolimerGroupId == 11)
                _nameBuilder.Append("т");
            if (cable.CoverPolimerGroupId == 10)
                _nameBuilder.Append("м");
            if (cable.CoverPolimerGroupId == 6)
                _nameBuilder.Append("нг(А)-LS");
            if (cable.CoverPolimerGroupId == 7)
                _nameBuilder.Append("нг(А)-LSLTx");

            var cableConductorArea = areaInSqrMm;
            namePart = CableCalculations.FormatConductorArea(cableConductorArea);
            _nameBuilder.Append($" {cable.ElementsCount}х2х{namePart}");

            return _nameBuilder.ToString();
        }

    }
}
