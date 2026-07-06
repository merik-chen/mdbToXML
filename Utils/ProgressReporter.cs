using System.Diagnostics;

namespace MdbToXml.Utils;

public class ProgressReporter
{
    private readonly bool _verbose;
    private readonly Stopwatch _stopwatch = new();
    private readonly object _lock = new();

    public ProgressReporter(bool verbose)
    {
        _verbose = verbose;
        _stopwatch.Start();
    }

    public void SetTotalTables(int totalTables)
    {
        // For possible future use to report total progress
    }

    public void Report(string tableName, int rowCount)
    {
        if (!_verbose) return;
        var elapsed = _stopwatch.Elapsed;
        var rate = rowCount / Math.Max(elapsed.TotalSeconds, 0.001);
        Console.Write($"\r  [{tableName}] {rowCount:N0} rows ({rate:N0} rows/sec)   ");
    }

    public void TableCompleted(string tableName, int totalRows, TimeSpan elapsed)
    {
        var rate = totalRows / Math.Max(elapsed.TotalSeconds, 0.001);
        lock (_lock)
        {
            Console.WriteLine($"  \u2713 {tableName}: {totalRows:N0} rows ({rate:N0} rows/sec, {elapsed.TotalSeconds:F1}s)");
        }
    }

    public void PrintSummary(int totalRows, int totalTables)
    {
        _stopwatch.Stop();
        var elapsed = _stopwatch.Elapsed;
        Console.WriteLine();
        Console.WriteLine($"  ======================================");
        Console.WriteLine($"  Export completed:");
        Console.WriteLine($"    Tables: {totalTables}");
        Console.WriteLine($"    Rows:   {totalRows:N0}");
        Console.WriteLine($"    Time:   {elapsed.TotalSeconds:F2}s");
        if (elapsed.TotalSeconds > 0)
            Console.WriteLine($"    Speed:  {totalRows / elapsed.TotalSeconds:N0} rows/sec");
        Console.WriteLine($"  ======================================");
    }
}
