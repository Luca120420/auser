using System;

class TestTimeConversion
{
    static void Main()
    {
        // Test the conversion logic
        double decimalValue = 0.354166666666667;
        
        double totalHours = decimalValue * 24.0;
        int hours = (int)totalHours;
        double decimalMinutes = totalHours - hours;
        int minutes = (int)(decimalMinutes * 60.0);
        
        string timeString = $"{hours:D2}:{minutes:D2}";
        
        Console.WriteLine($"Input: {decimalValue}");
        Console.WriteLine($"Total hours: {totalHours}");
        Console.WriteLine($"Hours: {hours}");
        Console.WriteLine($"Decimal minutes: {decimalMinutes}");
        Console.WriteLine($"Minutes: {minutes}");
        Console.WriteLine($"Output: {timeString}");
        
        // Test another value
        decimalValue = 0.375;
        totalHours = decimalValue * 24.0;
        hours = (int)totalHours;
        decimalMinutes = totalHours - hours;
        minutes = (int)(decimalMinutes * 60.0);
        timeString = $"{hours:D2}:{minutes:D2}";
        
        Console.WriteLine($"\nInput: {decimalValue}");
        Console.WriteLine($"Output: {timeString}");
    }
}
