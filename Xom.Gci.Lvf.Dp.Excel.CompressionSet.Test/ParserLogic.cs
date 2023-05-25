using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xom.Gci.Lvf.Dp.Helper.Common;
using Xom.Gci.Lvf.Dp.Helper.ParserJson;
using Newtonsoft.Json;

namespace Xom.Gci.Lvf.Dp.Excel.CompressionSet.Test
{
    
        public class ParserLogic
        {
            public static bool ExecuteHardnessParser(FileParserJson fileParserJson)
            {
                bool parserResult = false;
                try
                {
                    string fileParserResult = JsonConvert.SerializeObject(fileParserJson);
                    var strParameter = Base64.Base64Encode(fileParserResult);
                    string[] arg = new string[] { strParameter };
                    var result = FileProcessor.ParseFile(arg);
                    if (result.ToString() != "0")
                    {
                        parserResult = true;
                    }
                }
                catch (Exception ex)
                {
                    parserResult = false;
                }
                return parserResult;
            }

            public static bool CompareOutputResult(string outputJsonPath, string testDataJsonPath)
            {
                bool state = false;
                try
                {
                    string outputs = System.IO.File.ReadAllText(outputJsonPath);
                    FileParserJsonModel output = JsonConvert.DeserializeObject<FileParserJsonModel>(outputs);

                    string test = System.IO.File.ReadAllText(testDataJsonPath);
                    FileParserJsonModel Test = JsonConvert.DeserializeObject<FileParserJsonModel>(test);

                    var objJson = JsonConvert.SerializeObject(output);
                    var anotherJson = JsonConvert.SerializeObject(Test);

                    state = objJson == anotherJson;
                }
                catch
                {
                    state = false;
                }
                return state;
            }
        }


   
}
