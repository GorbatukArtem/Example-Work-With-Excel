using OfficeOpenXml;
using System.Diagnostics;

namespace EppPlus_Speed_Test.Source;

public static class Excel
{
    public static Stopwatch Stopwatch { get; set; }

    public static IEnumerable<object[]> CreateArray(int rowCount, int columnCount)
    {
        for (int rowIndex = 1; rowIndex <= rowCount; rowIndex++)
        {
            var newObjects = new object[columnCount];

            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                newObjects[columnIndex] = (columnIndex + 1) * rowIndex;
            }

            yield return newObjects;
        }
    }

    public static void Save(IEnumerable<object[]> array)
    {
        using (var package = new ExcelPackage())
        {
            using (var worksheet = package.Workbook.Worksheets.Add("EpPlus"))
            {
                worksheet.Cells[1, 1].LoadFromArrays(array);

                var folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\epplus\\";
                var filePath = $"{folderPath}{DateTime.Now:yyyymmdd_HHmmss}.xlsx";

                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);

                if (File.Exists(filePath))
                    File.Delete(filePath);

                var fileStream = File.Create(filePath);

                fileStream.Close();

                File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
        }
    }

    public static void StartClock()
    {
        Stopwatch = new Stopwatch();
        Stopwatch.Start();
    }

    public static void StopClock()
    {
        Stopwatch.Stop();
    }

    public static TimeSpan Elapsed()
    {
        return Stopwatch.Elapsed;
    }
}
