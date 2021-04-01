using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing.CableTitleBulders
{
    public interface ICableTitleBuilder
    {
        string GetCableTitle(Cable cable, InsulatedBillet mainBillet, Cables.Common.CableProperty? cableProperty, object parameter = null);
    }
}
