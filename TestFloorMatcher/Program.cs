using System;
using System.Linq;
using SpecCollector;

class Program
{
    static void Main()
    {
        string docPath = @"C:\Users\Lerik\YandexDisk\templates\Цех\Объекты\Амбер\Исходные\Generator.grb";
        
        var matcher = new FloorMatcher(docPath);
        
        string plane = "(5-1)-(5-6)";
        int floorFrom = 2, floorTo = 10;
        
        Console.WriteLine("=== GetTypicalFloorStats('" + plane + "', " + floorFrom + ", " + floorTo + ") ===");
        var stats = matcher.GetTypicalFloorStats(plane, floorFrom, floorTo);
        foreach (var kvp in stats)
        {
            Console.WriteLine("  Типовой этаж " + kvp.Key + ": " + kvp.Value + " раз(а)");
        }
        
        Console.WriteLine("\n=== CountTypicalFloorInRange ===");
        for (int target = 1; target <= 5; target++)
        {
            int count = matcher.CountTypicalFloorInRange(plane, floorFrom, floorTo, target);
            if (count > 0)
                Console.WriteLine("  Типовой этаж " + target + " в диапазоне [" + floorFrom + ".." + floorTo + "]: " + count + " раз(а)");
        }
        
        Console.WriteLine("\n=== Non-existent floor ===");
        Console.WriteLine("  Типовой этаж 99: " + matcher.CountTypicalFloorInRange(plane, 2, 10, 99) + " раз(а)");
        Console.WriteLine("  Non-existent plane: " + matcher.CountTypicalFloorInRange("nonexistent", 2, 10, 3) + " раз(а)");
    }
}
