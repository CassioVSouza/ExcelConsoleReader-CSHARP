using OfficeOpenXml;
using ReaderAPI;

namespace ExcelReader
{
    public class Program
    {
        static void Main()
        {
            Console.WriteLine("Welcome to the Excel reader! Please put your file path, Example: \"C:/Excels/Archive.xlsx\"  ");
            string filePath = Console.ReadLine();

            var data = ReadExcel.ReadExcelFile(filePath);
            var List = ReadExcel.TurnDataIntoList(data);

            int Choice = 0;
            do
            {
                Console.WriteLine("What do you want to do with the values:\n1 - Average\n2 - Biggest Number\n3 - Smallest Number\n4 - Show Numbers\n5 - Sum All\n6 - Show Numbers in grade\n7 - Show First Numbers");
                Choice = Convert.ToInt32(Console.ReadLine());
                switch (Choice)
                {
                    case 1:
                        ReadExcel.DoAverage(List);
                        break;
                    case 2:
                        ReadExcel.BiggestNumber(List);
                        break;
                    case 3:
                        ReadExcel.SmallestNumber(List);
                        break;
                    case 4:
                        ReadExcel.ShowNumber(List);
                        break;
                    case 5:
                        ReadExcel.SumNumbers(List);
                        break;
                    case 6:
                        ReadExcel.DisplayData(data);
                        break;
                    case 7:
                        ReadExcel.ShowFirstNumbers(List);
                        break;
                    default: Console.WriteLine("Insert a valid operation!");
                        break;
                }
                Console.WriteLine("\t\t");
            } while (Choice != 99);
          //  ReadExcel.DisplayData(data);
        }
    }
}

