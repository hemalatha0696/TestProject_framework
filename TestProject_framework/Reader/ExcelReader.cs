using ExcelDataReader;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject_framework.Reader
{
    public class ExcelReader
    {
        //String excelpath = @"C:\\Users\\al4104\\source\\repos\\qatoolwebsite\\qatoolwebsite\\TestData_ToolsQA\\Data_Reader.xlsx";

        ISheet sheet;
        XSSFWorkbook hssfwb;
        public int getRowCount(String sheetName)
        {
            sheet = hssfwb.GetSheet(sheetName);
            int rowCount = sheet.LastRowNum + 1;
            return rowCount;
        }

        public int getColumnCount(String sheetName)
        {
            sheet = hssfwb.GetSheet(sheetName); 
            XSSFRow rows = (XSSFRow)sheet.GetRow(0);
            int colCount = rows.LastCellNum + 1;    
            return colCount;
        }
        public static void Sheet_row_col(String sheetname)
        {
            String excelpath = @"C:\Users\al4104\source\repos\TestProject_framework\TestProject_framework\Reader\Data_Reader.xlsx";

            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(excelpath, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet(sheetname);
            DataFormatter df = new DataFormatter();
            //Create a row object to retrieve row at index 1

            //get all rows in the sheet
            int rowCount = sheet.LastRowNum - sheet.FirstRowNum;

            //iterate over all the row to print the data present in each cell.
            for (int i = 0; i <= rowCount; i++)
            {

                //get cell count in a row
                int cellcount = sheet.GetRow(i).LastCellNum;

                //iterate over each cell to print its value
                Console.WriteLine("Row" + i + " data is :");

                for (int j = 0; j < cellcount; j++)
                {
                    Console.WriteLine(sheet.GetRow(i).GetCell(j).StringCellValue + ",");
                }
                Console.WriteLine();
            }
        }
        //public static void WriteData(String sheetname, String RunMode, int row, int column)
        //{
        //    String excelpath = "C:\\Users\\al4104\\source\\repos\\qatoolwebsite\\qatoolwebsite\\TestData_ToolsQA\\DataSheet.xlsx";

        //    using (var fs = new FileStream(@excelpath, FileMode.Create, FileAccess.Write))
        //    {
        //        fs.Position = 0;
        //        IWorkbook hssfwb = new XSSFWorkbook(fs);
        //        ISheet sheet = hssfwb.GetSheet(sheetname);
        //        DataFormatter df = new DataFormatter();
        //        XSSFRow row2 = (XSSFRow)sheet.GetRow(row);
        //        sheet.GetRow(row).CreateCell(6).SetCellValue("Passed");
        //        hssfwb.Write(fs);
        //        fs.Close();
        //    }
        //}

        public static string ReadData(String sheetname, String RunMode, int row, int column)
        {
            XSSFWorkbook hssfwb;
            String excelpath = @"C:\Users\al4104\source\repos\TestProject_framework\TestProject_framework\Reader\Data_Reader.xlsx";

            try
            {
                using (FileStream file = new FileStream(excelpath, FileMode.Open, FileAccess.ReadWrite))
                {
                    file.Position = 0;

                    hssfwb = new XSSFWorkbook();
                    hssfwb.Write(file);
                }
                ISheet sheet = hssfwb.GetSheet(sheetname);
                DataFormatter df = new DataFormatter();

                XSSFRow rowRunmode = (XSSFRow)sheet.GetRow(row);
                XSSFCell cellRunMode = (XSSFCell)rowRunmode.GetCell(column);
                String valueofRunMode = df.FormatCellValue(cellRunMode);
                if (valueofRunMode == "Run Mode")
                {
                    //int sheetrowcount = ExcelReader.getRowCount("DataSheet");
                    for (int i = 3; i < sheet.LastRowNum; i++)
                    {
                        XSSFRow getrow = (XSSFRow)sheet.GetRow(i);

                        XSSFCell getcol = (XSSFCell)getrow.GetCell(0);
                        String RunModeValue = df.FormatCellValue(getcol);
                        if (RunMode == RunModeValue)
                        {
                            XSSFRow row2 = (XSSFRow)sheet.GetRow(row);
                            Console.WriteLine(row2);
                            XSSFCell cell = (XSSFCell)row2.GetCell(column);
                            Console.WriteLine(cell);
                            //As XSSFCell in NOPI cannot be converted to string we are using DataFormatter class 
                            String valueofcell = df.FormatCellValue(cell);

                            return valueofcell;

                        }
                        if (RunMode == RunModeValue)// && TestCaseName.Equals("Elementspage_TextBox"))
                        {
                            XSSFRow row2 = (XSSFRow)sheet.GetRow(row);
                            Console.WriteLine(row2);
                            XSSFCell cell = (XSSFCell)row2.GetCell(column);
                            Console.WriteLine(cell);
                            //As XSSFCell in NOPI cannot be converted to string we are using DataFormatter class 
                            String valueofcell = df.FormatCellValue(cell);

                            return valueofcell;

                        }
                        else
                            return null;

                    }
                }



                return null;
            }

            catch (Exception e)
            {
                throw;
            }
          
            }

    }
}
