using CableDataParsing.TableEntityes;

namespace CableDataParsing.NameBuilders
{
    public interface ICableNameBuilder<T> where T : CablePresenter
    {
        string GetCableName(T cable, InsulatedBilletPresenter insBillet, ConductorPresenter conductor, object parameter);
    }
}
