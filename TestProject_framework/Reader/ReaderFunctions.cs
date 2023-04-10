using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestProject_framework.Reader
{
    public class ReaderFunctions
    {
        private static IDictionary<string, IExcelDataReader> _cache;
        //private static FileStream stream;
        private static IExcelDataReader reader;
        static ReaderFunctions()
        {
            _cache = new Dictionary<string, IExcelDataReader>();
        }
        public static int GetTotalRows(string xlPath, string sheetName)
        {
            int count = 0;
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(xlPath, FileMode.Open, FileAccess.Read))
            {
                using (reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            count = count + 1;
                        }
                    } while (reader.NextResult()) ;
                }
            }
            return count;
        }

        public static object GetCellData(string xlPath, string sheetName, int row, int column)
        {
            List<List<object>> columns = new List<List<object>>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(xlPath, FileMode.Open, FileAccess.Read))
            {
                using (reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            if (i <= columns.Count)
                                columns.Add(new List<object>());
                            columns[i].Add(reader.GetValue(i));
                            //Console.WriteLine(reader.GetValue(i));
                        }
                    } while (reader.NextResult()) ;
                    // reader =  ExcelReaderFactory.CreateReader(stream);
                    // _cache.Add(sheetName, reader);
                }
                reader.Close();

                // return reader;
                //}

                //for (int i = 0; i < columns.Count; i++)
                //{
                //    for (int j = 0; j < columns[i].Count; j++)
                //    {
                //        Console.WriteLine("values of {0} at {1} is {2}", i, j, columns[i][j]);

                //    }
                //}
            }
            return reader;
        }

    }
}
