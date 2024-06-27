using System;
using System.Text.RegularExpressions;

// Shehu Yipmong Said BU/22C/IT/7533
public class Program
{
    public static void Main()
    {
        while (true)
        {
            Console.WriteLine("Enter a date mm-dd-yyyy or type 'exit' to quit:");
            string input = Console.ReadLine();

            if (input.ToLower() == "exit")
            {
                break;
            }

            try
            {
                string result = ReverseDateFormat(input);
                Console.WriteLine($"Converted date: {result}");
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error: {e.Message}");
            }
        }
    }

    public static string ReverseDateFormat(string dateInput)
    {
        if (string.IsNullOrWhiteSpace(dateInput))
        {
            throw new ArgumentException("Input date cannot be null or empty.");
        }

        // Define the regex pattern
        string pattern = @"^(?<mon>\d{1,2})-(?<day>\d{1,2})-(?<year>\d{2,4})$";
        TimeSpan timeout = TimeSpan.FromMilliseconds(500); // Set a timeout of 500 milliseconds

        try
        {
            // Match the input date with the regex pattern
            Match match = Regex.Match(dateInput, pattern, RegexOptions.None, timeout);

            if (match.Success)
            {
                // Extract the components
                string month = match.Groups["mon"].Value;
                string day = match.Groups["day"].Value;
                string year = match.Groups["year"].Value;

                // Normalize year to 4 digits
                if (year.Length == 2)
                {
                    year = "20" + year; // Assuming dates in 2000-2099 range
                }

                // Validate the date
                string normalizedDate = $"{month}-{day}-{year}";
                if (DateTime.TryParse(normalizedDate, out DateTime parsedDate))
                {
                    // Return the date in yyyy-MM-dd format
                    return parsedDate.ToString("yyyy-MM-dd");
                }
                else
                {
                    throw new ArgumentException("Invalid date format.");
                }
            }
            else
            {
                throw new ArgumentException("Invalid date format.");
            }
        }
        catch (RegexMatchTimeoutException)
        {
            throw new TimeoutException("The regex operation timed out.");
        }
    }
}
