using ExcelLibrary.SpreadSheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CustomMerge
{
    internal class Program
    {
        static List<int> positions = new List<int>();
        static int idPosition = 0;
        static List<Product> products = new List<Product>();
        static List<Product> productsWithoutDuplicate = new List<Product> { };
        static List<string> titles = new List<string>();

        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            if (!GetSettings())
            {
                Console.WriteLine("Press Enter to close");
                Console.ReadLine();
                return;
            }

            if (!CollectAllExcelData())
            {
                Console.WriteLine("Press Enter to close");
                Console.ReadLine();
                return;
            }

            foreach (Product product in products)
            {
                if (!ProductIsInNonDuplicateList(product))
                {
                    productsWithoutDuplicate.Add(product);
                }
                else
                {
                    Console.WriteLine("Merging " + product.ToString());
                    bool success = mergeProduct(product);
                    if (!success)
                    {
                        Console.ReadLine();
                        return;
                    }
                 
                }
            }

            createMergedExcel();

            foreach(Product product in productsWithoutDuplicate)
            {
                Console.WriteLine(product.data[0] + " " + product.data[1]);
            }

            Console.WriteLine("Finished!");
            Console.WriteLine("Press Enter to close");
            Console.ReadLine();
        }

        static bool GetSettings()
        {

            bool success = false;
            bool idFound = false;
            if (!File.Exists("Setting.txt"))
            {
                Console.WriteLine("Setting.txt not found");
                return success;
            }
            string line = File.ReadLines("Setting.txt").First();
            string[] column = line.Split(';');
            List<int> positionsOfSemicolon = new List<int>();
            int position = 0;
            foreach (string s in column)
            {
                if (s.ToLower().Equals("true"))
                {
                    positionsOfSemicolon.Add(position);
                }
                if (s.ToLower() == "id")
                {
                    idFound = true;
                }
                position++;
            }

            if (!idFound){
                Console.WriteLine("Setting.txt needs at least 1 ID");
                return success;
            }

            positions = positionsOfSemicolon;

            position = 0;
            foreach (string s in column)
            {
                if (s.ToLower().Equals("id"))
                {
                    idPosition = position;
                    success = true;
                    break;
                }
                position++;
            }

            return success;
        }
    
        static bool CollectAllExcelData()
        {
            DirectoryInfo dinfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            FileInfo[] Files = dinfo.GetFiles("*.xls");
            if (Files.Length < 1)
            {
                Console.WriteLine("no xls found");
                return false;
            }
            //Save titles
            FileInfo first = Files[0];
            var workbookT = Workbook.Load(first.Name);
            var worksheetT = workbookT.Worksheets[0]; // assuming only 1 worksheet
            var cellsT = worksheetT.Cells;
            int columnLengthT = cellsT.LastColIndex;
            for (int j = 0; j <= columnLengthT; j++)
            {
                titles.Add(cellsT[0, j].ToString());
            }

            foreach (FileInfo fi in Files)
            {
                var workbook = Workbook.Load(fi.Name);
                var worksheet = workbook.Worksheets[0]; // assuming only 1 worksheet
                var cells = worksheet.Cells;
                int columnLength = cells.LastColIndex;
                int rowLength = cells.LastRowIndex;
                
                for (int i = 0; i < rowLength; i++)
                {
                    if (isNullOrEmpty(worksheet.Cells[i,0].ToString()))
                    {
                        continue;
                    }

                    Product product = new Product();
                    List<string> data = new List<string>();
                    for (int j = 0; j <= columnLength; j++)
                    {
                        data.Add(worksheet.Cells[i+1, j].ToString());  
                    }
                    product.data = data;
                    products.Add(product);
                }
            }
            return true;
        }
    
        static bool ProductIsInNonDuplicateList(Product toCheck)
        {
            foreach(Product product in productsWithoutDuplicate)
            {
                if (product.data[idPosition].Equals(toCheck.data[idPosition]))
                {
                    return true;
                }
            }
            return false;
        }

        static bool mergeProduct(Product toMerge)
        {
            foreach(Product product in productsWithoutDuplicate)
            {
                if (product.data[idPosition] == toMerge.data[idPosition])
                {
                    Product newProduct = new Product();
                    List<string> newList = new List<string>();
                    newList = product.data;
                    foreach (int position in positions)
                    {
                        try
                        {
                            string oldValue1 = product.data[position];
                            string oldValue2 = toMerge.data[position];
                            if ((oldValue1 == null || oldValue1 == "") || (oldValue2 == null || oldValue2 == ""))
                            {
                                continue;
                            }
                            string newValue = int.Parse(oldValue1) + int.Parse(oldValue2) + "";
                            newList[position] = newValue;
                        }
                        catch (FormatException e)
                        {
                            Console.WriteLine(e.StackTrace);
                            Console.WriteLine(e.ToString());
                            continue;
                        }
                        catch (Exception e) {
                            Console.WriteLine(e.ToString());
                            continue;
                        }
                    }
                    newProduct.data = newList;
                    productsWithoutDuplicate.Remove(product);
                    productsWithoutDuplicate.Add(newProduct);
                    break;
                }
            }
            return true;
        }

        static void createMergedExcel()
        {
            DateTime now = DateTime.Now;
            string currentDate = now.ToString("yyyy'-'MM'-'dd' 'HH'-'mm'-'ss");
            string filename = "Merged List - " + currentDate+".xls";
            Workbook workbook = new Workbook();
            Worksheet worksheet = new Worksheet("First Sheet");

            //Creating empty cells to avoid error
            for (int i= 0; i < 100; i++)
            {
                worksheet.Cells[i, 0] = new Cell("");
            }

            //Creating titles
            int startPos = 0;
            foreach (string title in titles)
            {
                worksheet.Cells[0, startPos] = new Cell(title);
                startPos++;
            }

            //Insert data
            int startPosCol = 0;
            int startPosRow = 1;
            foreach(Product product in productsWithoutDuplicate)
            {
                foreach(string data in product.data)
                {
                    worksheet.Cells[startPosRow, startPosCol] = new Cell(product.data[startPosCol]);
                    startPosCol++;
                }
                startPosCol = 0;
                startPosRow++;
            }

            workbook.Worksheets.Add(worksheet);
            workbook.Save(filename);
        }

        static bool isNullOrEmpty(string str)
        {
            return str == null || str == "" || str.Length<1;
        }

    }

    class Product
    {
        public List<string> data = new List<string>();
        public override string ToString()
        {
            string str = "data: ";
            foreach(string s in data)
            {
                str += s+" ";
            }
            return str;
        }
    }

    class Setting
    {
        public int [] position { get; set; }
    }

}
