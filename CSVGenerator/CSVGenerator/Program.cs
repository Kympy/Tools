using Microsoft.Office.Interop.Excel;

public class CSVGenerator
{
    public static void Main()
    {
        string originPath = Directory.GetCurrentDirectory();
        string filePath = originPath;
        Console.WriteLine($"Current Directory is {filePath}");
        if (Directory.Exists(filePath) == false)
        {
            Console.WriteLine("No Directory");
            return;
        }
        string fileName = "asdf";
        filePath = $"{filePath}\\{fileName}.xlsx";
        if (File.Exists(filePath) == false)
        {
            Console.WriteLine("No File");
            return;
        }
        Console.WriteLine($"Load File Name : {fileName}..");
        Application app = new Application();
        try
        {
            Console.WriteLine("Open file..");
            Workbook workbook = app.Workbooks.Open(filePath);
            Worksheet sheet = workbook.Worksheets.Item[1] as Worksheet;
            sheet.SaveAs($"{originPath}\\{fileName}.csv", XlFileFormat.xlCSV);
            workbook.Close();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
            Console.WriteLine("Failed to open!");
        }
        app.Quit();
        return;
    }
}
