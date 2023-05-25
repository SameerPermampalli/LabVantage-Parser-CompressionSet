using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Xom.Gci.Lvf.Dp.Helper.ParserJson;

namespace Xom.Gci.Lvf.Dp.Excel.CompressionSet.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod()]
        public void BTEC_10175_BTEC_10261()
        {
            string folderPath = AppDomain.CurrentDomain.BaseDirectory;

            FileParserJson fileParserJson = new FileParserJson
            {
                ManifestPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\test.manifest.xml"),
                MappingPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\test.Mapping.txt"),
                TestDefinitionPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\test.ResultInput.json"),
                ResultsOutputPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\test.ResultOutPut.json"),
                AVTestDefinitionPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\av.ResultInput.json"),
                AVResultsOutputPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\av.ResultOutPut.json"),
                ResultsPath = Path.Combine(folderPath, "BTEC_10175_BTEC_10261"),
                ResultFilepaths = new Resultfilepath[]
                {

                    new Resultfilepath
                    {
                        Filepath= Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\BTEC_10175 and BTEC_10261 Compression Set  Data Sheet.xlsx")
                    }

                }
            };

            bool res = false;
            try
            {
                if (ParserLogic.ExecuteHardnessParser(fileParserJson))
                {
                    res = ParserLogic.CompareOutputResult(fileParserJson.ResultsOutputPath, Path.Combine(folderPath, "BTEC_10175_BTEC_10261\\test.json"));
                }


            }
            catch (Exception exp)
            {
                throw exp;
            }
            Assert.IsTrue(res);
        }
    }
}
