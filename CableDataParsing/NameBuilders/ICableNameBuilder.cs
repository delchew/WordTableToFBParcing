using CableDataParsing.TableEntityes;

namespace CableDataParsing.NameBuilders
{
    public interface ICableNameBuilder<T> where T : CablePresenter
    {
        string GetCableName(T cable, CableBilletPresenter insBillet, ConductorPresenter conductor, object parameter);
    }
}
