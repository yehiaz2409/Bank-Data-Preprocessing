using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;

class Program
{
    static void Main()
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var file = new FileInfo(@"D:\ABK Internship\Backend\Task 1\Task 1\YPE140240229.xlsx");
        string outputFilePath = @"D:\ABK Internship\Backend\Task 1\Task 1\TransformedData.txt";

        using (var package = new ExcelPackage(file))
        {

            var worksheet = package.Workbook.Worksheets[0];

            using (StreamWriter writer = new StreamWriter(outputFilePath))
            {
                // Write the header line
                string header = GenerateHeader();
                writer.WriteLine(header);
                int count = 0;
                decimal total = 0;
                // Iterate over the rows in the Excel worksheet
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Assuming row 1 is the header
                {
                    // Read data from the current row
                    string bankId = worksheet.Cells[row, 1].Text;
                    string date = worksheet.Cells[row, 2].Text;
                    string cardNumber = worksheet.Cells[row, 3].Text;
                    string transactionCode = worksheet.Cells[row, 4].Text;
                    string transactionCategory = worksheet.Cells[row, 5].Text;
                    string currency = worksheet.Cells[row, 6].Text;
                    string amount = worksheet.Cells[row, 7].Text;
                    string postingDescription = worksheet.Cells[row, 8].Text;

                    // Generate the transaction line
                    string transactionLine = GenerateTransactionLine(bankId, date, cardNumber, transactionCode, transactionCategory, currency, amount, postingDescription);
                    writer.WriteLine(transactionLine);
                    count++;
                    total += decimal.Parse(amount);
                }
                total *= 100;
                string totalTransactions = count.ToString().PadLeft(5, '0');
                string totalAmount = total.ToString().PadLeft(15, '0');
                string trailer = GenerateTrailer(totalTransactions, totalAmount);
                writer.WriteLine(trailer);
            }
        }

            Console.WriteLine($"File saved successfully as {outputFilePath}");
    }
    static string ConvertToCYYMMDD(DateTime date)
    {
        // Determine the century part
        string century = (date.Year / 100).ToString();

        // Extract the last two digits of the year
        string year = (date.Year % 100).ToString("D2");

        // Get the month and day
        string month = date.Month.ToString("D2");
        string day = date.Day.ToString("D2");

        // Combine to form the CYYMMDD format
        return $"{century}{year}{month}{day}";
    }

    static string GenerateHeader()
    {
        // Generate the header line as per the rules provided
        string recordId = "01";
        string processingDate = ConvertToCYYMMDD(DateTime.Now); // CYYMMDD format
        string space_1 = new string(' ', 13); // Empty 13 characters
        string totalTransactions = "00000";
        string totalAmount = "000000000000000";
        string participantId = "140";
        string accountNumber = new string(' ', 19); // Empty 19 characters
        string transactionDate = new string('0', 7); // "0000000"
        string space_2 = new string(' ', 39);
        string transaction_cat = "00";
        string space_3 = "    ";
        string zeros = new string('0', 30);

        return $"{recordId}{processingDate}{space_1}{totalTransactions}{totalAmount}{participantId}{accountNumber}{transactionDate}{space_2}{transaction_cat}{space_3}{zeros}";
    }
    static string GenerateTrailer(string totalTransactions, string totalAmount)
    {
        // Generate the header line as per the rules provided
        string recordId = "03";
        string processingDate = ConvertToCYYMMDD(DateTime.Now); // CYYMMDD format
        string space_1 = new string(' ', 13); // Empty 13 characters
        string participantId = "140";
        string accountNumber = new string(' ', 19); // Empty 19 characters
        string transactionDate = new string('0', 7); // "0000000"
        string space_2 = new string(' ', 39);
        string transaction_cat = "00";
        string space_3 = "    ";
        string zeros = new string('0', 30);

        return $"{recordId}{processingDate}{space_1}{totalTransactions}{totalAmount}{participantId}{accountNumber}{transactionDate}{space_2}{transaction_cat}{space_3}{zeros}";
    }
    static string GenerateTransactionLine(string bankId, string date, string cardNumber, string transactionCode, string transactionCategory, string currency, string amount, string postingDescription)
    {
        // Format the transaction line as per the rules provided
        string recordId = "02";
        string processingDate = ConvertToCYYMMDD(DateTime.Now); // CYYMMDD format
        string space_1 = new string(' ', 13); // Empty 13 characters
        string zeros_1 = new string('0', 5);
        string participantId = bankId.PadLeft(3, '0');
        string accountNumber = cardNumber.PadRight(19, ' '); // Card number with 19 characters
        string amountPart = (decimal.Parse(amount) * 100).ToString("000000000000000").PadLeft(15, '0'); // Amount with 15 digits
        string transactionDate = ConvertToCYYMMDD(DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture)); // CYYMMDD format
        string space_2 = new string(' ', 18);
        string postingDescriptionPart = postingDescription.PadRight(40, ' '); // Posting Description with 40 characters
        string transactionType = transactionCode.PadRight(2, ' '); // Transaction type
        string transactionCategoryPart = transactionCategory.PadLeft(2, '0'); // Transaction Category with 2 digits
        string debitCreditFlag = "C"; // Assume credit
        string currencyPart = currency.PadRight(3, ' '); // Currency with 3 characters
        string space_3 = new string(' ', 55);
        //string emptyFields = new string(' ', 65); // Empty spaces for remaining fields


        return $"{recordId}{processingDate}{space_1}{zeros_1}{amountPart}{participantId}{accountNumber}{transactionDate}{accountNumber}{space_2}{transactionType}{transactionCategoryPart}{debitCreditFlag}{currencyPart}{amountPart}{amountPart}{currencyPart}{space_3}{postingDescriptionPart}";
    }
}
