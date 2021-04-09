using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing.CableBulders
{
    public interface ICableTitleBuilder
    {
        string GetCableTitle(Cable cable, InsulatedBillet mainBillet, Cables.Common.CablePropertySet? cableProperty, object parameter = null);
    }
}
