using C1.C1Excel;
using log4net;
using log4net.Core;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xom.Gci.Lvf.Dp.Excel.NSpecimen.Lib.Model;
using Xom.Gci.Lvf.Dp.Excel.NSpecimen.Lib.Service;
using Xom.Gci.Lvf.Dp.Helper.Common;
using Xom.Gci.Lvf.Dp.Helper.Manifest;
using Xom.Gci.Lvf.Dp.Helper.ParserJson;
using static Xom.Gci.Lvf.Dp.Helper.ErrorHandling.ErrorLog;

namespace Xom.Gci.Lvf.Dp.Excel.CompressionSet
{

    public class FileProcessor
    {


        public static int ParseFile(string[] args)
        {
            int exitCode = 0;
            try
            {



                Log($"File parsing started...", Level.Info);
                Log(args[0], Level.Debug);

                ExcelFileParser excelParser = new ExcelFileParser();
                string resultJson = Base64.Base64Decode(args[0]);
                Resultfilepath matchedResultFile = new Resultfilepath();
                List<Resultfilepath> resultfilePaths = new List<Resultfilepath>();
                string excelFile = string.Empty;
                FileParserJson fileParserJson = Lvf.Dp.Excel.NSpecimen.Lib.Help.CommonHelp.ConvertJsonToClass<FileParserJson>(resultJson);
                string tempMapping = Path.Combine(Path.GetDirectoryName(fileParserJson.MappingPath), "temp", Path.GetFileName(fileParserJson.MappingPath));
                string mappingNew = Path.Combine(Path.GetDirectoryName(fileParserJson.MappingPath), string.Concat("temp_", Path.GetFileName(fileParserJson.MappingPath)));

                string tempExcelFile = Path.Combine(Path.GetDirectoryName(fileParserJson.MappingPath), "temp", "Result.xlsx");
                string tempFolder = Path.Combine(Path.GetDirectoryName(fileParserJson.MappingPath), "temp");
                string extartlocation = Path.Combine(Path.GetDirectoryName(fileParserJson.MappingPath), "temp");
                string requestItem = ManifestHelper.GetManifestFromXml(fileParserJson.ManifestPath).RequestItemID;
                string sampleID = ManifestHelper.GetManifestFromXml(fileParserJson.ManifestPath).SampleID;
                bool outPutGenerated = false;

                //check if package contain any .zip file
                if (!(fileParserJson.ResultFilepaths.Length > 0))
                {
                    Log(NoDataFile(), Level.Error, true);
                    return exitCode;
                }
                else
                {


                    resultfilePaths = fileParserJson.ResultFilepaths.Where(c => Path.GetFileName(c.Filepath).ToLower().Contains(".xls")).ToList();

                    if (resultfilePaths.Count == 0)
                    {
                        Log(NoDataFile(), Level.Error, true);
                        return exitCode;
                    }

                    int finalRow = 0, finalCol = 0, valueColumn = 0, valueRow = 0;
                    string sheetname = string.Empty;
                    string compressionValue = string.Empty;
                    bool sampleFound=false;

                    int maxrow = 0;
                    int maxcolumns = 0;

                    foreach (Resultfilepath filepath in resultfilePaths)
                    {
                        C1XLBook book = new C1XLBook();

                        try
                        {
                            book.Load(filepath.Filepath);
                            List<XLSheet> sheets = new List<XLSheet>();
                            foreach (XLSheet wsheet in book.Sheets)
                            {
                                sheets.Add(wsheet);
                            }
                               
                            foreach (XLSheet wsheet in  sheets)
                            {


                                Log($"Found Sheet with Request ID...", Level.Info);

                                maxrow = wsheet.Rows.Count;
                                maxcolumns = wsheet.Columns.Count;
                                XLCell cellValue = null;

                                for (int row = 0; row < maxrow; row++)
                                {
                                    for (int col = 0; col < maxcolumns; col++)
                                    {
                                        cellValue = wsheet.GetCell(row, col);

                                        if (cellValue != null && cellValue.Value != null)
                                        {

                                            if (cellValue.Value.ToString().Trim().ToLower().Contains(requestItem.ToLower()))
                                            {
                                                Log($"Found Request IDS at... row {row} col {col}", Level.Info);

                                                sampleFound = true;
                                                finalRow = row;
                                                finalCol = col;
                                                sheetname = wsheet.Name;
                                                valueColumn = finalCol + 7;
                                                valueRow = row + 1;


                                                #region create temp file

                                                // create file

                                                if (Directory.Exists(tempFolder))
                                                {
                                                    Directory.Delete(tempFolder, true);
                                                }
                                                Directory.CreateDirectory(tempFolder);                                                                                              

                                                File.Copy(filepath.Filepath, tempExcelFile);

                                                //end

                                            //    XLCell value = wsheet.GetCell(valueRow, valueColumn);
                                            //    string cellValueString = string.Empty;
                                             //   if (value != null && value.Value != null)
                                             //   {
                                               //     cellValueString = cellValue.Value.ToString().Trim();
                                               // }


                                                book.Sheets.Add("Results");
                                                XLSheet wsheetR = book.Sheets["Results"];

                                               // wsheet[valueRow-1, valueColumn].Value = "Compression Set";
                                               // wsheet[valueRow, valueColumn].Value = cellValueString;
                                               // compressionValue = cellValueString;
                                                book.Save(tempExcelFile);

                                                #endregion 
                                                break;

                                            }

                                        }

                                    }

                                    if (sampleFound)
                                    {
                                        break;
                                    }
                                }
                                if (sampleFound)
                                {
                                    break;
                                }

                            }

                            string micrometer = string.Empty;
                            string micrometerName = string.Empty;
                            string serialNo = string.Empty;
                            string specimenType = string.Empty;
                            string Time = string.Empty;
                            string Percent = string.Empty;
                            string Temp = string.Empty;
                            string OrderNO = string.Empty;
                            string Lubricatio = string.Empty;
                            string dateIn = string.Empty;
                            string dateOut = string.Empty;
                            string testMethod = string.Empty;
                            string molded = string.Empty;

                            if (sampleFound)
                            {
                                book.Load(tempExcelFile);

                                XLSheet sheet = book.Sheets[sheetname];

                                for (int row = 0; row < 9; row++)
                                {

                                    for (int col = finalCol; col < finalCol + 7; col++)
                                    {
                                        XLCell cellValue = null;
                                        cellValue = sheet.GetCell(row, col);

                                        if (cellValue != null && cellValue.Value != null)
                                        {
                                            
                                            if (cellValue.Value.ToString().Contains("Micrometer"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row + 1, col);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    micrometerName = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Serial No."))
                                            {
                                                //Serial No
                                                XLCell cellValue2 = sheet.GetCell(row, col+3);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    serialNo = cellValue2.Value.ToString();
                                                }
                                                // Specimen Type 2
                                                XLCell cellValue3 = sheet.GetCell(row, col + 5);
                                                if (cellValue3 != null && cellValue3.Value != null)
                                                {
                                                    specimenType = cellValue3.Value.ToString();
                                                }
                                                //test metod
                                                XLCell cellValue4 = sheet.GetCell(row+1, col + 5);
                                                if (cellValue4 != null && cellValue4.Value != null)
                                                {
                                                    testMethod = cellValue4.Value.ToString();
                                                }

                                            }
                                            if (cellValue.Value.ToString().Contains("Time:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col+1);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    Time = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Percent:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col+1);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    Percent = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Temp:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col+1);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    Temp = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Date IN:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    dateIn = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Date Out:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    dateOut = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Order No:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col+1);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    OrderNO = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Molded @"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    molded = cellValue2.Value.ToString();
                                                }
                                            }
                                            if (cellValue.Value.ToString().Contains("Lubrication:"))
                                            {
                                                XLCell cellValue2 = sheet.GetCell(row, col);
                                                if (cellValue2 != null && cellValue2.Value != null)
                                                {
                                                    Lubricatio = cellValue2.Value.ToString();
                                                }
                                            }
                                        }
                                    }
                                }


                                XLCell compcellValue = null;
                               
                                    compcellValue = sheet.GetCell(valueRow, valueColumn);

                                if (compcellValue != null && compcellValue.Value != null)
                                {
                                    compressionValue = compcellValue.Value?.ToString();
                                }

                                XLSheet resultSheet = book.Sheets["Results"];

                                resultSheet[0, 0].Value = "CompressionSet";
                                resultSheet[0, 1].Value = compressionValue;

                                resultSheet[11, 0].Value = "Micrometer";
                                resultSheet[11, 1].Value = micrometerName;

                                resultSheet[1, 0].Value = "Seial No.";
                                resultSheet[1, 1].Value = serialNo;

                                resultSheet[2, 0].Value = "Specimen";
                                resultSheet[2, 1].Value = specimenType;

                                resultSheet[3, 0].Value = "Time:";
                                resultSheet[3, 1].Value = Time.Replace("hr","").Trim();
                                resultSheet[3, 2].Value = "hr";

                                resultSheet[4, 0].Value = "Percent:";
                                resultSheet[4, 1].Value = Percent.Replace("%","").Trim();
                                resultSheet[4, 2].Value = "%";

                                resultSheet[5, 0].Value = "Temp:";
                                resultSheet[5, 1].Value = Temp;

                                resultSheet[6, 0].Value = "Order No:";
                                resultSheet[6, 1].Value = OrderNO;

                                resultSheet[7, 0].Value = "Lubrication:";
                                resultSheet[7, 1].Value = Lubricatio.Replace("Lubrication:", "").Trim(); ;

                                resultSheet[8, 0].Value = "Date IN:";
                                resultSheet[8, 1].Value = dateIn.Replace("Date IN:", "").Trim();

                                resultSheet[9, 0].Value = "Date Out:";
                                resultSheet[9, 1].Value = dateOut.Replace("Date Out:", "").Trim(); ;

                                resultSheet[10, 0].Value = "Test Method";
                                resultSheet[10, 1].Value = testMethod;

                                

                                book.Save(tempExcelFile);
                            }

                        }
                        catch (Exception exp)
                        {
                            Log(exp?.Message + " " + exp?.InnerException?.Message, Level.Error);
                        }

                    }

                    Log($"Created Excel", Level.Info);

                    // end tempfile
                    if(!sampleFound)
                    {
                        Log("Unable to find matching Request Item ID in file", Level.Info, true);
                        return 0;
                    }
                    
                    fileParserJson.ResultFilepaths = new Resultfilepath[]
                    {
                                               new Resultfilepath
                                               {
                                                   Filepath= tempExcelFile
                                               }
                                       };



                    //preare for processing excel file
                    Log("Started parsing Excel.", Level.Info);
                    string fileParserResult = JsonConvert.SerializeObject(fileParserJson);
                    var strParameter = Base64.Base64Encode(fileParserResult);
                    Log($"Executing Excel parser for {strParameter}.", Level.Info);

                    ResultModel result = new ResultModel();
                    result = excelParser.FileParser(strParameter);

                    if (result.Result)
                    {
                        if (DeleteTempFolder(Path.GetDirectoryName(tempExcelFile)))
                        {
                            Log($"Deleted Excel file {excelFile}", Level.Info);
                        }
                        Log("Parser executed successfully!", Level.Info, true);
                        exitCode = 1;
                        // break;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(result.Message) && result.Message.Contains("##ErrorCode##"))
                        {
                            if (!outPutGenerated)
                                Log(result.Message, Level.Error, true);
                            outPutGenerated = true;
                        }
                        else
                        {
                            if (!outPutGenerated)
                                Log(OtherError($"Failed to parser file: {result?.Message}"), Level.Error, true);
                            outPutGenerated = true;
                        }
                    }

                }

            }
            catch (Exception e)
            {
                if (e != null && ((!string.IsNullOrEmpty(e.Message) && (e.Message.Contains("##ErrorCode##"))) || e.InnerException != null && ((!string.IsNullOrEmpty(e.InnerException.Message) && (e.InnerException.Message.Contains("##ErrorCode##"))))))
                {
                    Log(e?.Message + " " + e?.InnerException?.Message, Level.Error, true);
                }
                else
                {
                    Log(OtherError(e?.Message + " " + e?.InnerException?.Message), Level.Error, true);
                }
            }
            return exitCode;
        }



        private static readonly log4net.ILog _iLog = LogManager.GetLogger(typeof(FileProcessor));
        public static bool DeleteTempFolder(string tempPath)
        {
            bool status = false;
            try
            {
                if (ConfigurationManager.AppSettings["DeleteTemp"] == null && Directory.Exists(tempPath))
                {
                    Directory.Delete(tempPath, true);
                }
                if (ConfigurationManager.AppSettings["DeleteTemp"].ToLower() == "yes")
                {
                    Directory.Delete(tempPath, true);
                    status = true;
                }
            }
            catch (Exception exp)
            {
                _iLog.Error(exp);
            }
            return status;
        }

    }

}
