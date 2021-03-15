using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CableDataParsing.MSWordTableParsers;
using Cables.Common;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using Microsoft.EntityFrameworkCore;
using WordObj = Microsoft.Office.Interop.Word;

namespace CableDataParsing
{
    public class KpsvevParser : CableParser
    {
        private StringBuilder _nameBuilder = new StringBuilder();

        public KpsvevParser(string connectionString, FileInfo mSWordFile) : base(connectionString, mSWordFile)
        { }

        public override int ParseDataToDatabase()
        {
            int recordsCount = 0;

            var app = new WordObj.Application { Visible = false };
            object fileName = _mSWordFile.FullName;

            try
            {
                app.Documents.Open(ref fileName);
                var document = app.ActiveDocument;
                var tables = document.Tables;
                if(tables.Count > 0)
                {
                    _wordTableParser = new WordTableParser
                    {
                        
                    };
                    List<TableCellData> tableData;

                    var cablePropertiesList = _dbContext.CableProperties.ToList();
                    var climaticModUHL = _dbContext.ClimaticMods.Find(3);
                    var redColor = _dbContext.Colors.Find(1);
                    var blackColor = _dbContext.Colors.Find(2);
                    var cableShortName = _dbContext.CableShortNames.Find(6); // КПСВ(Э)
                    var operatingVoltage = _dbContext.OperatingVoltages.Find(6); // 300В 50Гц, постоянка - до 500В
                    var TU2 = _dbContext.TechnicalConditions.Find(1);
                    var TU30 = _dbContext.TechnicalConditions.Find(35);

                    throw new NotImplementedException();

                }
                else
                    throw new Exception("Отсутствуют таблицы для парсинга в указанном Word файле!");
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                app.Quit();
            }

            return recordsCount;
        }

        public int ParseBillets()
        {
            var conuctor05 = _dbContext.Conductors.Find(17);
            var conuctor075 = _dbContext.Conductors.Find(18);
            var conuctor1 = _dbContext.Conductors.Find(19);
            var conuctor15 = _dbContext.Conductors.Find(20);
            var conuctor25 = _dbContext.Conductors.Find(21);

            var cableShortName = _dbContext.CableShortNames.Find(6); // КПСВ(Э)
            var operatingVoltage = _dbContext.OperatingVoltages.Find(6); // 300В 50Гц, постоянка - до 500В

            var polymerGroups = new List<PolymerGroup>
            {
                _dbContext.PolymerGroups.Find(1), // PVC
                _dbContext.PolymerGroups.Find(6), //PVC LS
                _dbContext.PolymerGroups.Find(7), //PVC LSLTx
                _dbContext.PolymerGroups.Find(11) //PVC term
            };

            List<InsulatedBillet> kpsvvBillets;

            foreach(var group in polymerGroups)
            {
                kpsvvBillets = new List<InsulatedBillet>
                {
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor05,
                        Diameter = 1.8m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor075,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.06m : 1.96m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor1,
                        Diameter = group.Id == 6 || group.Id == 7 ? 2.29m : 2.19m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor15,
                        Diameter = 2.66m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    },
                    new InsulatedBillet
                    {
                        CableShortName = cableShortName,
                        Conductor = conuctor25,
                        Diameter = 3.06m,
                        OperatingVoltage = operatingVoltage,
                        PolymerGroup = group
                    }
                };
                _dbContext.AddRange(kpsvvBillets);
            }
            return _dbContext.SaveChanges();
        }

        public override string GetCableTitle(Cable cable, params object[] cableParametres)
        {
            if (cableParametres != null)
                if (cableParametres[0] is Cables.Common.CableProperty cableProps)
                {
                    _nameBuilder.Clear();
                    _nameBuilder.Append("КПСВ");
                    if ((cableProps & Cables.Common.CableProperty.HasFoilShield) == Cables.Common.CableProperty.HasFoilShield)
                        _nameBuilder.Append("Э");
                    string namePart = string.Empty;
                    switch (cable.CoverPolymerGroup.Title)
                    {
                        case "PVC":
                        case "PVC Term":
                        case "PVC LS":
                        case "PVC Cold":
                            namePart = "В";
                            break;
                        case "PE":
                            namePart = "Пс";
                            break;
                    }
                    _nameBuilder.Append(namePart);

                    if ((cableProps & Cables.Common.CableProperty.HasArmourBraid) == Cables.Common.CableProperty.HasArmourBraid)
                    {
                        if ((cableProps & Cables.Common.CableProperty.HasArmourTube) == Cables.Common.CableProperty.HasArmourTube)
                            _nameBuilder.Append($"К{namePart}");
                        else
                            _nameBuilder.Append("КГ");
                    }
                    if ((cableProps & Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube) ==
                        (Cables.Common.CableProperty.HasArmourTape | Cables.Common.CableProperty.HasArmourTube))
                    {
                        _nameBuilder.Append($"Б{namePart}");
                    }

                    if (cable.CoverPolymerGroup.Title == "PVC Term")
                        _nameBuilder.Append("т");
                    if (cable.CoverPolymerGroup.Title == "PVC Cold")
                        _nameBuilder.Append("м");
                    if (cable.CoverPolymerGroup.Title == "PVC LS")
                        _nameBuilder.Append("нг(А)-LS");

                    var cableConductorArea = cable.ListCableBillets.First().Billet.Conductor.AreaInSqrMm;
                    namePart = CableCalculations.FormatConductorArea((double)cableConductorArea);
                    _nameBuilder.Append($" {cable.ElementsCount}х2х{namePart}");

                    return _nameBuilder.ToString();
                }
            throw new Exception("В метод не передан параметр Cables.Common.CableProperty или параметр не соответствует этому типу данных!");
        }
    }
}
