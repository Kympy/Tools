using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace TableLoader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }
        private const string TrunkName = "Trunk";
        private void FindFolder(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (fb.ShowDialog() == DialogResult.OK)
            {
                pathBox.Text = fb.SelectedPath;
                int index = fb.SelectedPath.IndexOf(TrunkName);
                if (index >= 0)
                {
                    string subPath = fb.SelectedPath.Substring(0, index + TrunkName.Length);
                    savePathBox.Text = subPath;
                    savePathBox.Text += "\\Client\\Assets\\Resources_moved\\Table";
                    classPathBox.Text = subPath;
                    classPathBox.Text += "\\Client\\Assets\\Scripts\\Table";
                }
            }
        }
        private void pathBox_TextChanged(object sender, EventArgs e)
        {
            int index = pathBox.Text.IndexOf(TrunkName);
            if (index >= 0)
            {
                string subPath = pathBox.Text.Substring(0, index + TrunkName.Length);
                savePathBox.Text = subPath;
                savePathBox.Text += "\\Client\\Assets\\Resources_moved\\Table";
                classPathBox.Text = subPath;
                classPathBox.Text += "\\Client\\Assets\\Scripts\\Table";
            }
        }
        private void fileFindButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            of.Title = "파일 경로 선택";
            of.Filter = "엑셀 파일 (*.xlsx)|*.xlsx;|모든 파일 (*.*)|*.*";

            DialogResult dr = of.ShowDialog();

            if (dr == DialogResult.OK)
            {
                fileTextBox.Text = of.FileName;
                int index = fileTextBox.Text.IndexOf(TrunkName);
                if (index >= 0)
                {
                    string subPath = fileTextBox.Text.Substring(0, index + TrunkName.Length);
                    savePathBox.Text = subPath;
                    savePathBox.Text += "\\Client\\Assets\\Resources_moved\\Table";
                    classPathBox.Text = subPath;
                    classPathBox.Text += "\\Client\\Assets\\Scripts\\Table";
                }
            }
        }
        private void fileTextBox_TextChanged(object sender, EventArgs e)
        {
            int index = fileTextBox.Text.IndexOf(TrunkName);
            if (index >= 0)
            {
                string subPath = fileTextBox.Text.Substring(0, index + TrunkName.Length);
                savePathBox.Text = subPath;
                savePathBox.Text += "\\Client\\Assets\\Resources_moved\\Table";
                classPathBox.Text = subPath;
                classPathBox.Text += "\\Client\\Assets\\Scripts\\Table";
            }
        }
        private void SetSavePathButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (fb.ShowDialog() == DialogResult.OK)
            {
                savePathBox.Text = fb.SelectedPath;
            }
        }

        private void GenerateAllButton_Click(object sender, EventArgs e)
        {
            GenerateAll(pathBox.Text);
        }
        private void GenerateFileButton_Click(object sender, EventArgs e)
        {
            GenerateAll(fileTextBox.Text, false);
        }
        // JSON 저장 경로
        public string savePath_Json;
        // 스크립트 저장 경로
        public string savePath_Script;
        public readonly string excelType = "xlsx";
        public Microsoft.Office.Interop.Excel.Application app;
        public Workbook workbook;
        public Worksheet worksheet;
        private void GenerateAll(string argPath, bool isAll = true)
        {
            try
            {
                logBox.Clear();
                if (isAll == true && Directory.Exists(argPath) == false)
                {
                    Log($"폴더를 찾을 수 없습니다. : \"{argPath}\"\n");
                    return;
                }
                if (Directory.Exists(savePathBox.Text) == false ||
                    Directory.Exists(classPathBox.Text) == false)
                {
                    Log($"저장 경로가 설정되지 않았거나 폴더가 존재하지 않습니다.\nJson : \"{savePathBox.Text}\"\nClass : \"{classPathBox.Text}\"\n");
                    return;
                }
                savePath_Json = savePathBox.Text;
                savePath_Script = classPathBox.Text;
                if (isAll == true)
                {
                    // 정해진 경로안의 파일들을 모두 가져옴 / Find all files in my directory
                    Log("폴더 내에 존재하는 모든 엑셀 파일 로드 중..\n");
                    List<string> filePath = new List<string>();
                    filePath.AddRange(Directory.GetFiles(argPath));
                    // 파일들의 갯수가 0이면 종료 / When file count is zero, return.
                    if (filePath.Count == 0)
                    {
                        Log("폴더 내에 파일이 없습니다.\n");
                        return;
                    }
                    Log("-------------------------------------------\n");
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
                        Log($"[{count}] : {filePath[i]}\n"); // If xlsx, write log.
                        count++;
                    }
                    Log("....................................\n");
                    count = 1;
                    for (int i = 0; i < filePath.Count; i++)
                    {
                        if (filePath[i].IndexOf("~$") != -1)
                        {
                            Log($"[{count}] 임시 엑셀 파일 제거 : {filePath[i]}\n"); // If xlsx, write log.
                            filePath.Remove(filePath[i]);
                            count++;
                            i--;
                            continue;
                        }
                    }
                    Log("-------------------------------------------\n");
                    Log($"최종 파일 리스트 : {filePath.Count} 개\n");
                    count = 1;
                    for (int i = 0; i < filePath.Count; i++)
                    {
                        Log($"[{count}] : {filePath[i]}\n");
                    }
                    Log("-------------------------------------------\n");
                    for (int i = 0; i < filePath.Count; i++)
                    {
                        if (Make(filePath[i]) == false)
                        {
                            Log($"생성 실패 -> {filePath[i]}\n");
                            Log("종료합니다.\n");
                            return;
                        }
                    }
                }
                else // 단일 파일 생성
                {
                    if (Make(argPath) == false)
                    {
                        Log($"생성 실패 -> {argPath}\n");
                        Log("종료합니다.\n");
                        return;
                    }
                }
                
                Log("엑셀 닫는중...\n");
                Log("성공적으로 모든 과정이 종료되었습니다.\n");
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
        private int StartCol = 1;
        private int StartRow = 3;
        private bool Make(string argFilePath)
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(argFilePath);
                worksheet = workbook.Worksheets.Item[1] as Worksheet;
                Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

                if (range.Cells[StartRow, StartCol] == null || (range.Cells[StartRow, StartCol] as Range).Value2 == null)
                {
                    Log("시작 셀(3, 1) 이 비어 있습니다.\n");
                    return false;
                }
                Log($"행 갯수 : {range.Rows.Count}, ");
                Log($"열 갯수 : {range.Columns.Count}\n");
                string[] names = new string[range.Columns.Count];
                string[] types = new string[range.Columns.Count];
                string[,] datas = new string[range.Rows.Count - 2, range.Columns.Count];
                // Get name / type cells
                Log("데이터 타입과 이름을 가져오는 중...\n");
                for (int column = StartCol; column <= range.Columns.Count; column++)
                {
                    types[column - StartCol] = (range.Cells[StartRow, column] as Range).Value2.ToString();
                    Log($"{types[column - StartCol]} : ");
                    names[column - StartCol] = (range.Cells[StartRow + 1, column] as Range).Value2.ToString();
                    Log($"{names[column - StartCol]}\n");
                }
                // Get data cells
                Log("데이터 읽는 중...\n");
                for (int column = StartCol; column <= range.Columns.Count; column++)
                {
                    for (int row = StartRow + 2; row <= range.Rows.Count; row++)
                    {
                        Log($"{(range.Cells[row, column] as Range).Value2}\n");
                        datas[row - (StartRow + 2), column - (StartCol)] = (range.Cells[row, column] as Range).Value2.ToString();
                    }
                }
                Log("데이터 읽기 완료.\n");
                Log("-------------------------------------------\n");
                // Split file name : ex -> 'testDocs' + '.xlsx'
                string[] fileName = workbook.Name.Split('.');
                // Make new save path
                if (Directory.Exists(savePath_Json) == false)
                {
                    Directory.CreateDirectory(savePath_Json);
                }
                string targetPath = $"{savePath_Json}\\{fileName[0]}.json";
                workbook.Close();
                app.Quit();
        
                Log($"JSON 파일 생성 중...\n");
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
                for (int j = 0; j < datas.GetLength(0) - 2; j++) // From 0 to row count
                {
                    writer.WriteLine("\t{");
                    for (int i = 0; i < names.Length; i++)
                    {
                        isArray = false;
                        writer.Write($"\t\t\"{names[i]}\": ");
                        if (string.Compare(types[i], "string") == 0)
                        {
                            writer.Write($"\"{datas[j, i]}\"");
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
                        else if (string.Compare(types[i], "float[]") == 0)
                        {
                            writer.Write("[");
                            writer.Write(datas[j, i]);
                            writer.WriteLine("]");
                            isArray = true;
                        }
                        else if (string.Compare(types[i], "bool[]") == 0)
                        {
                            writer.Write("[");
                            writer.Write(datas[j, i]);
                            writer.WriteLine("]");
                            isArray = true;
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
                    if (j == datas.GetLength(0) - 2 - 1)
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
                Log("JSON 저장 완료.\n");
        
                Log("C# 클래스 작성중...\n");
                if (Directory.Exists(savePath_Script) == false)
                {
                    Directory.CreateDirectory(savePath_Script);
                }
                targetPath = $"{savePath_Script}\\{fileName[0]}.cs";
                writer = File.CreateText(targetPath);
        
                writer.WriteLine("// Auto Created by Json writer program. create by DragonGate Table Loader.");
                writer.WriteLine();
                writer.WriteLine("[System.Serializable]");
                writer.WriteLine($"public class {fileName[0]} : Data");
                writer.WriteLine("{");
                for (int i = 1; i < datas.GetLength(1); i++) // ID 칼럼은 스킵한다.
                {
                    writer.WriteLine($"\tpublic {types[i]} {names[i]};");
                }
                writer.WriteLine("}");
                writer.WriteLine("[System.Serializable]");
                writer.WriteLine($"public class {fileName[0]}Data : TableBase<{fileName[0]}> {{ }}");
                writer.Close();
                Log("C# 클래스 저장 완료.\n");
                return true;
            }
            catch(Exception ex)
            {
                Log($"에러 발생 : {ex}\n");
                return false;
            }
        }
        private void Log(string message)
        {
            logBox.AppendText(message);
            logBox.ScrollToCaret();
        }

        private void OpenJsonButton_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(savePathBox.Text) == true)
            {
                Process.Start("Explorer.exe", savePathBox.Text);
            }
            else
            {
                Log("잘못된 경로입니다.\n");
            }
        }

        private void OpenClassButton_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(classPathBox.Text) == true)
            {
                Process.Start("Explorer.exe", classPathBox.Text);
            }
            else
            {
                Log("잘못된 경로입니다.\n");
            }
        }

        private void classFolderFindButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (fb.ShowDialog() == DialogResult.OK)
            {
                classPathBox.Text = fb.SelectedPath;
            }
        }
    }
}
