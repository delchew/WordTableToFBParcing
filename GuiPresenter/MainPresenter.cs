using CableDataParsing;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GuiPresenter
{
    public class MainPresenter
    {
        private readonly IView _view;
        private readonly IMessageService _messageService;
        private Dictionary<string, Func<CableParser>> _cableTypesDict;
        private string[] _dbConnectionsNames;
        private string _cableName;
        private string _connectionString;
        private IConfigurationRoot _connectionStringConfig;
        private Stopwatch _stopwatch;

        public MainPresenter(IView view, IMessageService messageService)
        {
            Initialize();

            _view = view;
            _messageService = messageService;

            var cablesNames = _cableTypesDict.Keys.ToArray();
            _view.SetCablesNames(cablesNames);
            _view.SetDBConnectionsNames(_dbConnectionsNames);

            GetConnectionString();

            _view.CableNameChanged += View_CableNameChanged;
            _view.TableParseStarted += View_TableParseStarted;
            _view.DBConnectionNameChanged += View_DBConnectionNameChanged;

            _stopwatch = new Stopwatch();
        }

        private void Initialize()
        {
            _cableTypesDict = new Dictionary<string, Func<CableParser>>
            {
                { "КПСВ(Э)", () => new KpsvevParser(_connectionString, _view.MSWordFile) },
                { "КУНРС", () => new KunrsParser(_connectionString, _view.MSWordFile) },
                { "СКАБ", () => new SkabParser(_connectionString, _view.MSWordFile) },
                { "КЭВ(Э)В, КЭРс(Э)", () => new Kevv_KerspParser(_connectionString, _view.MSWordFile) },
                { "КИП", () => new KipParser(_connectionString, _view.MSWordFile) }
            };

            _dbConnectionsNames = new string[]
            {
                "JobConnection",
                "TestJobConnection",
                "HomePCTestConnection"
            };
        }

        private void GetConnectionString()
        {
            var builder = new ConfigurationBuilder();
            var jsonDir = Directory.GetCurrentDirectory();
            // установка пути к текущему каталогу
            builder.SetBasePath(jsonDir);
            // получаем конфигурацию из файла appsettings.json
            builder.AddJsonFile("appsettings.json");
            // создаем конфигурацию
            _connectionStringConfig = builder.Build();
            // возвращаем из метода строку подключения
            _connectionString = _connectionStringConfig.GetConnectionString(_view.DBConnectionName);
        }

        private void View_DBConnectionNameChanged(string obj)
        {
            _connectionString = _connectionStringConfig.GetConnectionString(_view.DBConnectionName);
        }

        private void View_CableNameChanged(string cableName)
        {
            _cableName = cableName;
        }

        private async void View_TableParseStarted()
        {
            try
            {
                if (!FileExists(_view.MSWordFile))
                    return;
                if(string.IsNullOrEmpty(_cableName) || string.IsNullOrEmpty(_connectionString))
                {
                    _messageService.ShowExclamation($"Не выбрано соединение или марка кабеля!");
                    return;
                }
                using (var parser = _cableTypesDict[_cableName]?.Invoke())
                {
                    if (parser != null)
                    {
                        parser.ParseReport += Parser_ParseReport;

                        _stopwatch.Reset();
                        _stopwatch.Start();
                        int recordsCount = await Task<int>.Factory.StartNew(parser.ParseDataToDatabase);
                        _stopwatch.Stop();
                        var timeSpan = _stopwatch.Elapsed;
                        var elapsedTime = string.Format("{0:00}:{1:00}:{2:00}", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
                        _messageService.ShowMessage($"Успешно! Число записей, занесённых в базу: {recordsCount}{Environment.NewLine}Времени прошло: {elapsedTime}");
                    }
                    else
                    {
                        _messageService.ShowError("Не задан парсер для этого типа кабеля!");
                    }
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

        private void Parser_ParseReport(double completedPersentage)
        {
            _view.UpdateProgress(completedPersentage);
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
