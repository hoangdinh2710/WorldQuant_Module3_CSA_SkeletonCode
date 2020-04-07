using Microsoft.Office.Interop.Excel;
using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WorldQuant_Module3_CSA_SkeletonCode
{
    class Program
    {
        static Excel.Workbook workbook;
        static Excel.Application app;

        static void Main(string[] args)
        {
            app = new Excel.Application();
            app.Visible = true;
            try
            {
                workbook = app.Workbooks.Open("property_pricing.xlsx", ReadOnly: false);
            }
            catch
            {
                SetUp();
            }

            var input = "";
            while (input != "x")
            {
                PrintMenu();
                input = Console.ReadLine();
                try
                {
                    var option = int.Parse(input);
                    switch (option)
                    {
                        case 1:
                            try
                            {
                                Console.Write("Enter the size: ");
                                var size = float.Parse(Console.ReadLine());
                                Console.Write("Enter the suburb: ");
                                var suburb = Console.ReadLine();
                                Console.Write("Enter the city: ");
                                var city = Console.ReadLine();
                                Console.Write("Enter the market value: ");
                                var value = float.Parse(Console.ReadLine());

                                AddPropertyToWorksheet(size, suburb, city, value);

                                Console.WriteLine("Add data to worksheet!");
                            }
                            catch
                            {
                                Console.WriteLine("Error: couldn't parse input");
                            }
                            break;
                        case 2:
                            Console.WriteLine("Mean price: " + CalculateMean());
                            break;
                        case 3:
                            Console.WriteLine("Price variance: " + CalculateVariance());
                            break;
                        case 4:
                            Console.WriteLine("Minimum price: " + CalculateMinimum());
                            break;
                        case 5:
                            Console.WriteLine("Maximum price: " + CalculateMaximum());
                            break;
                        default:
                            break;
                    }
                } 
                catch 
                {
                   
                }
            }

            // save before exiting
            workbook.Save();
            workbook.Close();
            app.Quit();
        }

        static void PrintMenu()
        {
            Console.WriteLine();
            Console.WriteLine("Select an option (1, 2, 3, 4, 5) " +
                              "or enter 'x' to quit...");
            Console.WriteLine("1: Add Property");
            Console.WriteLine("2: Calculate Mean");
            Console.WriteLine("3: Calculate Variance");
            Console.WriteLine("4: Calculate Minimum");
            Console.WriteLine("5: Calculate Maximum");
            Console.WriteLine();
        }

        static void SetUp()
        {
            // TODO: Implement this method 
            Excel.Workbook newWorkbook = app.Workbooks.Add();
            newWorkbook.Title = "property_pricing";
            
            Excel._Worksheet currentWorksheet = app.ActiveSheet;
            // Write Header
            currentWorksheet.Cells[1, "A"] = "ID";
            currentWorksheet.Cells[1, "B"] = "Size";
            currentWorksheet.Cells[1, "C"] = "Sub Urban";
            currentWorksheet.Cells[1, "D"] = "City";
            currentWorksheet.Cells[1, "E"] = "Market Value";
        }

        static void AddPropertyToWorksheet(float size, string suburb, string city, float value)
        {
            // TODO: Implement this method
            Excel._Worksheet currentWorksheet = app.ActiveSheet;
            // Get last row 
            Excel.Range last = currentWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = currentWorksheet.get_Range("A1", last);

            int lastRow = last.Row+1;
            // Add to lastRow
            currentWorksheet.Cells[lastRow, "A"] = lastRow -1 ;
            currentWorksheet.Cells[lastRow, "B"] = size;
            currentWorksheet.Cells[lastRow, "C"] = suburb;
            currentWorksheet.Cells[lastRow, "D"] = city;
            currentWorksheet.Cells[lastRow, "E"] = value;
        }

        static float CalculateMean()
        {
            // TODO: Implement this method
            Excel._Worksheet currentWorksheet = app.ActiveSheet;
            // Get last row 
            Excel.Range last = currentWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = currentWorksheet.get_Range("A1", last);

            int lastRow = last.Row;

            int total = 0;

            for (int i = 1; i < lastRow; i++)
            {
                var value = currentWorksheet.Cells[i+1, "E"].value;
                total += value;
            }

            var mean = total / (lastRow-1);

            return (float)mean;
        }

        static float CalculateVariance()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateMinimum()
        {
            // TODO: Implement this method
            return 0.0f;
        }

        static float CalculateMaximum()
        {
            // TODO: Implement this method
            return 0.0f;
        }
    }
}
