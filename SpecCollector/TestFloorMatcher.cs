using System;
using SpecCollector;

class TestFloorMatcher
{
    static void Main()
    {
        string docPath = @"C:\Users\Lerik\YandexDisk\templates\Цех\Объекты\Амбер\Исходные\Generator.grb";
        
        var matcher = new FloorMatcher(docPath);
        Console.WriteLine($"FilePath: {matcher.FilePath}");
        Console.WriteLine($"Available planes: {string.Join(", ", matcher._data.Keys)}");
        
        string plane = "(5-1)-(5-6)";
        if (matcher._data.ContainsKey(plane))
        {
            Console.WriteLine($"\nSheet '{plane}' data:");
            var sheet = matcher._data[plane];
            foreach (var kvp in sheet.OrderBy(x => x.Key))
            {
                Console.WriteLine($"  Этаж {kvp.Key} -> Типовой этаж {kvp.Value}");
            }
        }
        else
        {
            Console.WriteLine($"\nPlane '{plane}' not found!");
        }
        
        Console.WriteLine("\n--- GetTypicalFloorStats('(5-1)-(5-6)', 2, 10) ---");
        var stats = matcher.GetTypicalFloorStats("(5-1)-(5-6)", 2, 10);
        foreach (var kvp in stats)
        {
            Console.WriteLine($"  Типовой этаж {kvp.Key}: {kvp.Value} раз(а)");
        }
    }
}
