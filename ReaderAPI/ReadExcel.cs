using OfficeOpenXml;

namespace ReaderAPI
{
    public class ReadExcel
    {
        public static object[,] ReadExcelFile(string filePath)
        {
            using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                object[,] data = new object[end.Row, end.Column];

                for (int row = start.Row; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        data[row - 1, col - 1] = worksheet.Cells[row, col].Value;
                    }
                }

                return data;
            }
        }

        public static List<int> TurnDataIntoList(object[,] data)
        {
            List<int> Numbers = new List<int>();
            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    Numbers.Add(Convert.ToInt32(data[row, col]));
                }
            }

            return Numbers;
        }
        public static void DoAverage(List<int> Numbers)
        {
            double average = 0;
            double FinalAverage = 0;

            foreach (int Number in Numbers) 
            {
                average += Number;
            }

            FinalAverage = average / Numbers.Count();
            Console.WriteLine("The Average is: " + FinalAverage);
        }

        public static void BiggestNumber(List<int> Numbers)
        {
            double Biggest = Numbers[0];
            foreach (int Number in Numbers)
            {
                if(Number > Biggest) Biggest = Number;
            }
            Console.WriteLine("The Biggest number is: " + Biggest);
        }

        public static void SmallestNumber(List<int> Numbers)
        {
            double Smallest = Numbers[0];
            foreach (int Number in Numbers)
            {
                if (Number < Smallest) Smallest = Number;
            }
            Console.WriteLine("The Smallest number is: " + Smallest);
        }

        public static void ShowNumber(List<int> Numbers)
        {
            Console.WriteLine("The numbers are: ");
            foreach (int Number in Numbers)
            {
                Console.WriteLine(Number);
            }
        }

        public static void SumNumbers(List<int> Numbers)
        {
            double Sum = 0;
            foreach (int Number in Numbers)
            {
                Sum += Number;
            }
            Console.WriteLine("The numbers added together are: " + Sum);
        }

        public static void ShowFirstNumbers(List<int> Numbers)
        {
            Console.WriteLine("The first numbers are: ");
            for(int i = 0; i < 10 && i < Numbers.Count; i++)
            {
                Console.WriteLine(Numbers[i]);
            }
        }


        public static void DisplayData(object[,] data)
        {
            for (int row = 0; row < data.GetLength(0); row++)
            {
                Console.WriteLine("The numbers in grade are: ");
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    Console.Write(data[row, col] + "\t");
                }
                Console.WriteLine();
            }
        }
    }
}

