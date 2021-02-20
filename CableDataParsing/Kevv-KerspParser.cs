using CableDataParsing.MSWordTableParsers;
using System;
using System.Collections.Generic;
using System.IO;
using WordObj = Microsoft.Office.Interop.Word;
using CablesDatabaseEFCoreFirebird;
using CablesDatabaseEFCoreFirebird.Entities;
using System.Linq;

namespace CableDataParsing
{
    public class Kevv_KerspParser : ICableDataParcer
    {
        public event Action<int, int> ParseReport;

        private WordTableParser _wordTableParser;
        private FileInfo _mSWordFile;

        public Kevv_KerspParser(FileInfo mSWordFile)
        {
            _mSWordFile = mSWordFile;
        }

        public int ParseDataToDatabase()
        {
            int recordsCount = 0;

            //var app = new WordObj.Application { Visible = false };
            //object fileName = _mSWordFile.FullName;

            using (var dbContext = new CablesContext())
            {
                try
                {
                    var pvcBillets = dbContext.InsulatedBillets.ToList();
                    var rubberBillets = dbContext.InsulatedBillets.Where(b => b.CableShortName.ShortName.ToLower() == "кэрс").ToList();

                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }

            return recordsCount;
        }
    }
}
