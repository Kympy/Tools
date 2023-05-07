using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using UnityEngine;
using UnityEditor;
namespace JSON_Generator
{
    public class JsonGenerator
    {
        // 자동화를 할 엑셀 파일들이 있는 디렉토리 경로 / The directory where excel files exist.
        public static string directoryPath;
        // 저장 경로 / Save path where the file saved.
        public static string savePath_Json;
        public static string savePath_Script;
        public const string excelType = "xlsx";
        public static Microsoft.Office.Interop.Excel.Application app;
        public static Workbook workbook;
        public static Worksheet worksheet;
        private static void Main(string[] argPath)
        {
            if (argPath.Length != 3)
            {
                Console.WriteLine("Argument is not valid.");
                return;
            }
            directoryPath = argPath[0];
            savePath_Json = argPath[1];
            savePath_Script = argPath[2];
            try
            {
                if (Directory.Exists(directoryPath) == false)
                {
                    Console.WriteLine($"Cannot find the directory : {directoryPath}");
                    return;
                }
                // 정해진 경로안의 파일들을 모두 가져옴 / Find all files in my directory
                List<string> filePath = new List<string>();
                Console.WriteLine("Finding all files(.xlsx) in the directory...");
                filePath.AddRange(Directory.GetFiles(directoryPath));
                // 파일들의 갯수가 0이면 종료 / When file count is zero, return.
                if (filePath.Count == 0)
                {
                    Console.WriteLine("There's no files in this directory.");
                    return;
                }
                Console.WriteLine("-------------------------------------------");
                int count = 1;
                // 추가된 파일 중 xlsx가 아닌 것들을 제거 / Delete non .xlsx files in list.
                for (int i = 0; i < filePath.Count; i++)
                {
                    string type = filePath[i].Substring(filePath[i].Length - 4);
                    if (string.Compare(type, excelType) != 0)
                    {
                        filePath.Remove(filePath[i]); // Delete and continue.
                        i--;
                        continue;
                    }
                    Console.WriteLine($"[{count}] : {filePath[i]}"); // If xlsx, write log.
                    count++;
                }
                Console.WriteLine("....................................");
                count = 1;
                for (int i = 0; i < filePath.Count; i++)
                {
                    if (filePath[i].IndexOf("~$") != -1)
                    {
                        filePath.Remove(filePath[i]);
                        i--;
                        continue;
                    }
                    else
                    {
                        Console.WriteLine($"[{count}] Delete : {filePath[i]}"); // If xlsx, write log.
                        count++;
                    }
                }
                Console.WriteLine("-------------------------------------------");
                Console.WriteLine($"--Final List : File Count -> ({filePath.Count}) --");
                for (int i = 0; i < filePath.Count; i++)
                {
                    Console.WriteLine($"[{count}] : {filePath[i]}");
                }
                Console.WriteLine("-------------------------------------------");

                for (int i = 0; i < filePath.Count; i++)
                {
                    MakeJson(filePath[i]);
                }
                Console.WriteLine("Quit Application...");
                Console.WriteLine("Quit");
                return;
            }
            finally
            {
                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }
                if (workbook != null)
                {
                    Marshal.ReleaseComObject(workbook);
                }
                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                }
                GC.Collect();
            }
        }

        private static void MakeJson(string path)
        {
            app = new Microsoft.Office.Interop.Excel.Application();
            workbook = app.Workbooks.Open(path);
            worksheet = workbook.Worksheets.Item[1];
            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
            string[] names = new string[range.Columns.Count];
            string[] types = new string[range.Columns.Count];
            string[,] datas = new string[range.Rows.Count - 2, range.Columns.Count];
            // Get name / type cells
            Console.WriteLine("Get name and types...");
            Debug.Log("Get name and types...");
            for (int column = 1; column <= range.Columns.Count; column++)
            {
                names[column - 1] = (string)(range.Cells[1, column]).Text;
                Console.Write($"{names[column - 1]} : ");
                types[column - 1] = (string)(range.Cells[2, column]).Text;
                Console.WriteLine(types[column - 1]);
            }
            // Get data cells
            Console.WriteLine("Get datas...");
            for (int column = 1; column <= range.Columns.Count; column++)
            {
                for (int row = 3; row <= range.Rows.Count; row++)
                {
                    Console.WriteLine($"{(string)(range.Cells[row, column]).Text}");
                    datas[row - 3, column - 1] = (string)(range.Cells[row, column]).Text;
                }
            }
            Console.WriteLine("Successfully get data.");
            Console.WriteLine("-------------------------------------------");
            // Split file name : ex -> 'testDocs' + '.xlsx'
            string[] fileName = workbook.Name.Split('.');
            // Make new save path
            string targetPath = $"{savePath_Json}\\{fileName[0]}.json";
            Console.WriteLine($"Create new json file : {fileName[0]}.json");
            workbook.Close();
            app.Quit();

            Console.WriteLine("Write Json file...");
            StreamWriter writer = File.CreateText(targetPath);

            bool CheckLastComma(int i, int last)
            {
                if (i == last)
                {
                    return true;
                }
                return false;
            }
            bool isArray = false;
            writer.WriteLine("[");
            for (int j = 0; j < datas.GetLength(0); j++) // From 0 to row count
            {
                writer.WriteLine("\t{");
                for (int i = 0; i < names.Length; i++)
                {
                    isArray = false;
                    writer.Write($"\t\t\"{names[i]}\": ");
                    if (string.Compare(types[i], "string") == 0)
                    {
                        writer.Write($"{datas[j, i]}");
                    }
                    else if (string.Compare(types[i], "int") == 0)
                    {
                        writer.Write($"{datas[j, i]}");
                    }
                    else if (string.Compare(types[i], "int[]") == 0)
                    {
                        writer.Write("[");
                        writer.Write(datas[j, i]);
                        writer.WriteLine("]");
                        isArray = true;
                    }
                    else if (string.Compare(types[i], "bool") == 0)
                    {
                        writer.Write($"{datas[j, i]}");
                    }
                    else if (string.Compare(types[i], "float") == 0)
                    {
                        writer.Write($"{datas[j, i]}");
                    }
                    else
                    {
                        writer.Write($"// {datas[j, i]} Cannot define the type!!");
                    }
                    // 마지막이거나 배열이면 , 생략
                    if (CheckLastComma(i, names.Length - 1) || isArray == true) // Last or array
                    {
                        if (isArray == false)
                        {
                            writer.WriteLine();
                        }
                    }
                    else // 마지막이 아니고 배열이 아닐 때
                    {
                        writer.WriteLine(",");
                    }
                }
                // If last, don't write comma
                if (j == datas.GetLength(0) - 1)
                {
                    writer.WriteLine("\t}");
                }
                else
                {
                    writer.WriteLine("\t},");
                }
            }
            writer.Write("]");
            writer.Close();
            Console.WriteLine("Finished saving Json file.");
            Console.WriteLine("-------------------------------------------");

            Console.WriteLine("Writing C# script...");
            targetPath = $"{savePath_Script}\\{fileName[0]}.cs";
            writer = File.CreateText(targetPath);

            writer.WriteLine("// Auto Created by Json writer program. create by Kwon yong moon.");
            writer.WriteLine();
            writer.WriteLine($"public class {fileName[0]}");
            writer.WriteLine("{");
            for (int i = 0; i < datas.GetLength(1); i++)
            {
                writer.WriteLine($"\tpublic {types[i]} {names[i]};");
            }
            writer.Write("}");
            writer.Close();
            Console.WriteLine("Finished saving C# script file.");
        }
    }
}
