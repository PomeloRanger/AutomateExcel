using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomatingExcel
{
    class Program
    {
        /*
        * ========================================================================================
        * This is how to "validate" if you are on the right column
        * 1st: Check for 8 digits
        * 2nd: "Validate" if you are on the right column(Some columns have 8 digits but if it start with a specific 3 digit it's not a card number)(Check ValidateCardInitial)
        * 3rd : Check if you can convert the number to Hex (If you can't then well it's incorrect)
        * =========================================================================================
        * Separating
        * 1st : Separating it into two parts
        * 2nd : If the first 3 digit starts with 42D split into 3 digit / If it starts with other digit split into 4 digit
        * 3rd : Take the remainder digit and convert it into decimal
        */
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter your Excel file path");
            string excelFilePath = Console.ReadLine();

            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(excelFilePath))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets[1];

                foreach (var items in ws.Cells)
                {
                    if (Validate(items.Text))
                    {
                        string[] cardNumber;
                        cardNumber = Seperate(items.Text);

                        Console.WriteLine(items.Text + "\t" + cardNumber[0] + "\t" + cardNumber[1] + "\t");
                    }
                }
            }
        }

        private static bool Validate(string itemText)
        {
            const int cardLength = 8; 

            return ((itemText.Length == cardLength)
            && (ValidateCardInitial(itemText))
            && (int.TryParse(itemText, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int temp)));
        }

        private static bool ValidateCardInitial(string value)
        {
            List<string> validInitialValues = new List<string> {"A10", "AXI", "277"};

            string truncValue = value.Substring(0, 3);
            return (!validInitialValues.Contains(truncValue));
        }

        private static string[] Seperate(string value)
        {
            string[] cardNumber = new string[2];

            //Check if character begins with 42D
            if (value.Substring(0, 3) == "42D")
            {
                cardNumber[0] = value.Substring(0, 3);

                if ((int.TryParse(value.Substring(3, value.Length - 3), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int hexSecondPart)))
                {
                    cardNumber[1] = hexSecondPart.ToString();
                }
            }
            else
            {
                cardNumber[0] = value.Substring(0, 4);

                if ((int.TryParse(value.Substring(4, value.Length - 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int hexSecondPart)))
                {
                    cardNumber[1] = hexSecondPart.ToString();
                }
            }
            return cardNumber;
        }
    }
}
