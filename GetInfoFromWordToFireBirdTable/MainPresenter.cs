using GetInfoFromWordToFireBirdTable.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GetInfoFromWordToFireBirdTable
{
    public class MainPresenter
    {
        private readonly IView _view;
        private readonly IMessageService _messageService;
        private readonly Dictionary<string, Func<ICableDataParcer>> _cableTypesDict;
        private string _cableName;
        public MainPresenter(IView view, IMessageService messageService)
        {
            _cableTypesDict = new Dictionary<string, Func<ICableDataParcer>>
            {
                {"КУНРС", () => new KunrsParser(_view.MSWordFile) },
                {"СКАБ", () => new SkabParser(_view.MSWordFile) }
            };

            _view = view;
            _messageService = messageService;

            var cablesNames = _cableTypesDict.Keys.ToArray();
            _view.SetCablesNames(cablesNames);

            _view.CableNameChanged += View_CableNameChanged;
            _view.TableParseStarted += View_TableParseStarted;
        }

        private void View_CableNameChanged(string cableName)
        {
            _cableName = cableName;
        }

        private async void View_TableParseStarted()
        {
            try
            {
                if (!FileExists(_view.FBDatabaseFile) || !FileExists(_view.MSWordFile))
                    return;
                if(string.IsNullOrEmpty(_cableName))
                {
                    _messageService.ShowExclamation($"Марка кабеля не выбрана!");
                    return;
                }
                var parser = _cableTypesDict[_cableName]?.Invoke();
                if (parser != null)
                {
                    parser.ParseReport += Parser_ParseReport;

                    int recordsCount = await Task<int>.Factory.StartNew(parser.ParseDataToDatabase);

                    _messageService.ShowMessage($"Успешно! Число записей, занесённых в базу: {recordsCount}");
                }
                else
                {
                    _messageService.ShowError("Не задан парсер для этого типа кабеля!");
                }
            }
            catch (Exception ex)
            {
                _messageService.ShowError($"Parse failed. Exception message: {ex.Message}" );
            }
            finally
            {
                _view.ParseFinishReport();
            }
        }

        private void Parser_ParseReport(int parseOperationsCount, int completedOperationsCount)
        {
            _view.UpdateProgress(parseOperationsCount, completedOperationsCount);
        }

        private bool FileExists(FileInfo file)
        {
            if (!file.Exists)
            {
                _messageService.ShowError($"Отсутствует файл {file.FullName}. Проверьте правильность пути к файлу!");
                return false;
            }
            return true;
        }
    }
}
