using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using CableDataParsing.MSWordTableParsers;
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

        public int ParseDataToDatabase()
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
            throw new NotImplementedException();
        }

        public override string ParceDataToDatabase()
        {
            throw new NotImplementedException();
        }

        protected override string GetCableTitle(Cable cable, params object[] cableParametres)  //TODO
        {
            if (cableParametres != null)
                if (cableParametres[0] is Cables.Common.CableProperty cableProps)
                {
                    _nameBuilder.Clear();
                    _nameBuilder.Append("КПСВ");
                    if ((cableProps & Cables.Common.CableProperty.HasFoilShield) == Cables.Common.CableProperty.HasFoilShield)
                        _nameBuilder.Append("Э");
                    string namePart;
                    switch (cable.CoverPolymerGroup.Title)
                    {
                        case "PVC": namePart = "В";
                            break;
                        case "PVC LS": namePart = "Внг(А)-LS";
                            break;
                    }


                    return _nameBuilder.ToString();
                }
            throw new Exception("В метод не передан параметр Cables.Common.CableProperty или параметр не соответствует этому типу данынх!");
        }
    }
}
