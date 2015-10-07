using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GetExcelColumnValues
{
    class PrintExcelColumnValue
    {
        static void Main(string[] args)
        {
            PrintExcelColumn();
        }
        static void PrintExcelColumn()
        {
            Console.WriteLine("************* Print Excel Column Number ********************");
            Console.WriteLine("1. Numeric Input Type (Column Number)");
            Console.WriteLine("2. Aplpabetic Input Type (Column Name)");
            Console.WriteLine("3. Exit");
            Console.WriteLine("Enter Proper Option 1 for Number & 2 for Name");

            int inputType = 0;
            int printType = 0;


            try
            {
                inputType = Convert.ToInt32(Console.ReadLine());
                if (inputType == 3)
                    Environment.Exit(1);
                Console.WriteLine("1. Print a range");
                Console.WriteLine("2. Print specific value");
                Console.WriteLine("Enter Proper Option 1 for Range & 2 for Specific number");
                printType = Convert.ToInt32(Console.ReadLine());
                switch (inputType)
                {
                    case 1:
                        switch (printType)
                        {
                            case 1:
                                Console.WriteLine("Enter From value:");
                                int from = Convert.ToInt32(Console.ReadLine());
                                Console.WriteLine("Enter To value:");
                                int to = Convert.ToInt32(Console.ReadLine());
                                PrintExcelColumnName(from, to);
                                PrintExcelColumn();
                                return;
                            case 2:
                                Console.WriteLine("Enter Column number:");
                                int value = Convert.ToInt32(Console.ReadLine());
                                string getConvertedValue = ConvertExcelColumnNumberToName(value);
                                Console.WriteLine("Column Name of " + value.ToString() + " is= " + getConvertedValue);
                                PrintExcelColumn();
                                return;
                        }


                        return;
                    case 2:
                        switch (printType)
                        {
                            case 1:
                                Console.WriteLine("Enter From value:");
                                string from = (Console.ReadLine());
                                Console.WriteLine("Enter To value:");
                                string to = (Console.ReadLine());
                                PrintExcelColumnNumber(from, to);
                                PrintExcelColumn();
                                return;
                            case 2:
                                Console.WriteLine("Enter Column Name:");
                                string value = (Console.ReadLine());
                                int getConvertedValue = ConvertExcelColumnNameToNumber(value);
                                Console.WriteLine("Column number of " + value + " is= " + getConvertedValue.ToString());
                                PrintExcelColumn();
                                return;
                        }
                        return;

                    default:
                        Console.WriteLine("Input option in not valid.");
                        Console.ReadLine();
                        return;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Input value in not valid.");
                Console.ReadLine();
            }
        }
        static void PrintExcelColumnName(int from, int to)
        {
            for (; from <= to; from++)
            {
                Console.WriteLine(ConvertExcelColumnNumberToName(from));
            }
        }
        static void PrintExcelColumnNumber(string from, string to)
        {
            int tempFrom = ConvertExcelColumnNameToNumber(from);
            int tempTo = ConvertExcelColumnNameToNumber(to);
            PrintExcelColumnName(tempFrom, tempTo);

        }
        static int ConvertExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            //Convert Column Name into Upper Case
            columnName = columnName.ToUpperInvariant();

            int NumberCalci = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                NumberCalci = NumberCalci * 26;
                NumberCalci = NumberCalci + (columnName[i] - 'A' + 1);
            }

            return NumberCalci;
        }
        static string ConvertExcelColumnNumberToName(int columnNumber)
        {
            if (columnNumber == null) throw new ArgumentNullException("columnNumber");

            string setColumnName = String.Empty;
            int tempRemainder = 0;

            while (columnNumber > 0)
            {
                tempRemainder = (columnNumber - 1) % 26;
                setColumnName = Convert.ToChar(65 + tempRemainder).ToString() + setColumnName;
                columnNumber = (int)((columnNumber - tempRemainder) / 26);
            }

            return setColumnName;
        }
    }
}
