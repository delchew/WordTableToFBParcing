using GetInfoFromWordToFireBirdTable.CableEntityes;

namespace GetInfoFromWordToFireBirdTable.Common.CableNameBuilders
{
    interface ICableNameBuilder<T> where T: Cable
    {
        string GetCableName(T cable);
    }
}
