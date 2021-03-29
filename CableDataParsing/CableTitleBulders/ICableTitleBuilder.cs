using CablesDatabaseEFCoreFirebird.Entities;

namespace CableDataParsing.CableTitleBulders
{
    public interface ICableTitleBuilder
    {
        string GetCableTitle(Cable cable, Cables.Common.CableProperty? cableProperty);
    }
}
