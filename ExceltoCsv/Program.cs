using System;
using System.Collections.Generic;
using ReadWriteCsv;
using Aspose.Cells;
using System.IO;

namespace Excel
{
    class Program
    {
        static void Main(string[] args)
        {

            //string dirPath = "/home/suraj/Documents/";
            Aspose.Cells.Workbook workbook = new Workbook("ProductList.xlsx");
            Aspose.Cells.Worksheet sheet = workbook.Worksheets[0];
            int row1 = sheet.Cells.Rows.Count;
            int column = sheet.Cells.Columns.Count;

            bool Valueisgreaterthan = false, ErrorValue = false, ValueCheck = false;
            int increment = 0, count = 0, ErrorPoint = 0;
            CsvRow rowF = new CsvRow();
            object[] ExcelHeaderData = new object[15]
            {"PID","Product Id","Mfr-Name","Vendor_PN","MfrPN","Vendor_PN","Cost","Coo","Short Description","Long Description",
            "UPC","UOM","Sale-Start-Date","Sale-End-Date","Sales-Price"};
            object[] CsvHeaderData = new object[15]
            {"PID","Product Id","Mfr P/N","Mfr Name","Vendor Name","Vendor P/N","Price","COO",
            "Short Description","Long Description","UPC","UOM","Sale Start Date","Sale End Date","Sale Price"};
            void WriteTest()
            {
                // Write sample data to CSV file
                using (CsvFileWriter writer = new CsvFileWriter("ProductList.csv"))
                {
                    for (int i = 0; i < row1; i++)
                    {
                        CsvRow row = new CsvRow();

                        for (int j = 0; j < 15; j++)
                        {
                            // Console.WriteLine(sheet.Cells.GetCell(i, j)?.Value);
                            if (j == 8)
                            {
                                ErrorPoint = j;
                                ErrorValue = true;
                                ++j;
                            }
                            if (i <= 10000)
                            {

                                row.Add(String.Format("{0} ", sheet.Cells.GetCell(i, j)?.Value));

                            }
                            else
                            {
                                Valueisgreaterthan = true;
                            }
                        }
                        writer.WriteRow(row);

                    }
                }
                if (Valueisgreaterthan == true)
                {

                    using (CsvFileWriter writer = new CsvFileWriter("ProductList" + increment + ".csv"))
                    {
                        for (int i = 10001; i < row1; i++)
                        {
                            CsvRow row = new CsvRow();
                            for (int j = 0; j < 15; j++)
                            {
                                if (j == 8)
                                {
                                    ErrorPoint = j;
                                    ErrorValue = true;
                                    ++j;
                                }
                                if (count == 0)
                                {
                                    i = 0;
                                    row.Add(String.Format("{0} ", sheet.Cells.GetCell(i, j)?.Value));

                                }
                                if (i < 20000)
                                {
                                    row.Add(String.Format("{0} ", sheet.Cells.GetCell(i, j)?.Value));
                                }
                                else
                                {
                                    ValueCheck = true;
                                    ++increment;
                                }
                            }
                            count++;
                            writer.WriteRow(row);
                        }
                    }
                    count = 0;
                }
                if (ValueCheck == true)
                {

                    using (CsvFileWriter writer = new CsvFileWriter("ProductList" + increment + ".csv"))
                    {
                        for (int i = 20001; i < row1; i++)
                        {
                            CsvRow row = new CsvRow();
                            for (int j = 0; j < 15; j++)
                            {
                                if (j == 8)
                                {
                                    ErrorPoint = j;
                                    ErrorValue = true;
                                    ++j;
                                }
                                if (count == 0)
                                {
                                    i = 0;
                                    row.Add(String.Format("{0} ", sheet.Cells.GetCell(i, j)?.Value));
                                }
                                else
                                {
                                    row.Add(String.Format("{0} ", sheet.Cells.GetCell(i, j)?.Value));
                                }

                            }
                            count++;
                            writer.WriteRow(row);
                        }
                    }
                }
                if (ErrorValue == true)
                {
                    using (CsvFileWriter writer = new CsvFileWriter("error.xlsx"))
                    {

                        for (int i = 0; i < row1; i++)
                        {
                            CsvRow row = new CsvRow();
                            for (int j = ErrorPoint; j < ErrorPoint + 1; j++)

                            {
                                row.Add(String.Format("{0} ", sheet.Cells.GetCell(i, j)?.Value));
                            }
                            writer.WriteRow(row);
                        }
                    }
                }
            }
            WriteTest();
            Console.WriteLine("Conveted into CSV file Successfully !!!!\nThanks for using this app.");
        }
    }
}
