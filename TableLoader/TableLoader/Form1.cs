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
            //pivotPathBox.Text = pivotFolderName;

            if (File.Exists(Directory.GetCurrentDirectory() + "\\lastConfig") == false) return;

            StreamReader sr = new StreamReader(Directory.GetCurrentDirectory() + "\\lastConfig");
            saveJsonPathBox.Text = sr.ReadLine();
            saveClassPathBox.Text = sr.ReadLine();
            sr.Close();
        }

        private string pivotFolderName = "";

        private void FindFolder(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (fb.ShowDialog() == DialogResult.OK)
            {
                pathBox.Text = fb.SelectedPath;
                if (string.IsNullOrEmpty(pivotFolderName) == true)
                {
                    saveJsonPathBox.Text = $"{pathBox.Text}\\Generated";
                    saveClassPathBox.Text = saveJsonPathBox.Text;
                }
                else
                {
                    saveJsonPathBox.Text = $"{pivotFolderName}\\Generated";
                    saveClassPathBox.Text = saveJsonPathBox.Text;
                }
            }
        }

        private void pathBox_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(pathBox.Text) == true) return;

            if (string.IsNullOrEmpty(pivotFolderName) == true)
            {
                saveJsonPathBox.Text = $"{pathBox.Text}\\Generated";
                saveClassPathBox.Text = saveJsonPathBox.Text;
            }
            else
            {
                saveJsonPathBox.Text = $"{pivotFolderName}\\Generated";
                saveClassPathBox.Text = saveJsonPathBox.Text;
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

                if (string.IsNullOrEmpty(fileTextBox.Text) == false) return;
                if (fileTextBox.Text.Contains(".xlsx") == false) return;

                int fileNameStartIndex = fileTextBox.Text.IndexOf(of.SafeFileName);
                if (string.IsNullOrEmpty(pivotFolderName) == true)
                {
                    saveJsonPathBox.Text = fileTextBox.Text.Substring(0, fileNameStartIndex);
                    saveJsonPathBox.Text += "Generated";
                    saveClassPathBox.Text = saveJsonPathBox.Text;
                }
                else
                {
                    saveJsonPathBox.Text = $"{pivotFolderName}\\Generated";
                    saveClassPathBox.Text = saveJsonPathBox.Text;
                }
            }
        }

        private void fileTextBox_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(fileTextBox.Text) == false) return;
            if (fileTextBox.Text.Contains(".xlsx") == false) return;

            string[] temp = fileTextBox.Text.Split('\\');
            string fileName = temp[temp.Length - 1];

            int fileNameStartIndex = fileTextBox.Text.IndexOf(fileName);
            if (string.IsNullOrEmpty(pivotFolderName) == true)
            {
                saveJsonPathBox.Text = fileTextBox.Text.Substring(0, fileNameStartIndex);
                saveJsonPathBox.Text += "Generated";
                saveClassPathBox.Text = saveJsonPathBox.Text;
            }
            else
            {
                saveJsonPathBox.Text = $"{pivotFolderName}\\Generated";
                saveClassPathBox.Text = saveJsonPathBox.Text;
            }
            //int index = fileTextBox.Text.IndexOf(pivotFolderName);
            //         if (index >= 0)
            //         {
            //             string subPath = fileTextBox.Text.Substring(0, index);
            //             savePathBox.Text = subPath;
            //             savePathBox.Text += "\\Client\\Assets\\Resources_moved\\Table";
            //             classPathBox.Text = subPath;
            //             classPathBox.Text += "\\Client\\Assets\\Scripts\\Table";
            //         }
        }

        private void SetSavePathButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (fb.ShowDialog() == DialogResult.OK)
            {
                saveJsonPathBox.Text = fb.SelectedPath;
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
            logBox.Clear();
            if (isAll == true && Directory.Exists(argPath) == false)
            {
                Log($"폴더를 찾을 수 없습니다. : \"{argPath}\"\n");
                return;
            }

            if (string.IsNullOrEmpty(saveJsonPathBox.Text) == true ||
                string.IsNullOrEmpty(saveClassPathBox.Text) == true)
            {
                Log(
                    $"저장 경로가 설정되지 않았습니다.\nJson : \"{saveJsonPathBox.Text}\"\nClass : \"{saveClassPathBox.Text}\"\n");
                return;
            }

            if (Directory.Exists(saveJsonPathBox.Text) == false)
            {
                Directory.CreateDirectory(saveJsonPathBox.Text);
            }

            if (Directory.Exists(saveClassPathBox.Text) == false)
            {
                Directory.CreateDirectory(saveClassPathBox.Text);
            }

            savePath_Json = saveJsonPathBox.Text;
            savePath_Script = saveClassPathBox.Text;
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
                        workbook.Close();
                        app.Quit();
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
                    workbook.Close();
                    app.Quit();
                    return;
                }
            }

            Log("엑셀 닫는중...\n");
            Log("성공적으로 모든 과정이 종료되었습니다.\n");
            workbook.Close();
            app.Quit();
            return;
        }

        private void Release()
        {
            if (worksheet != null)
            {
                Marshal.ReleaseComObject(worksheet);
                worksheet = null;
            }

            if (workbook != null)
            {
                Marshal.ReleaseComObject(workbook);
                workbook = null;
            }

            if (app != null)
            {
                Marshal.ReleaseComObject(app);
                app = null;
            }

            GC.Collect();
            GC.SuppressFinalize(this);
        }

        private int StartCol;
        private int StartRow;

        private bool Make(string argFilePath)
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(argFilePath);
                worksheet = workbook.Worksheets.Item[1] as Worksheet;
                //Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

                StartRow = -1;
                // "Id" 로 시작하는 칼럼을 찾아 아래로 내려가면서 행을 찾는다.
                Range lastCell = (worksheet.Cells[1, "A"] as Range).Cells;
                for (int i = 0; i < 10; i++)
                {
                    // 빈 셀일 경우 다음으로
                    if (lastCell.Value2 == null)
                    {
                        lastCell = lastCell.get_End(XlDirection.xlDown).Cells;
                        continue;
                    }

                    // ID 가 아닐 경우 다음으로
                    if (string.Compare(lastCell.Value2.ToString(), "Id") != 0)
                    {
                        Log($"Pass : {lastCell.Value2.ToString()}\n");
                        lastCell = lastCell.get_End(XlDirection.xlDown).Cells;
                        continue;
                    }
                    else
                    {
                        StartRow = lastCell.Row;
                        break;
                    }
                }

                if (StartRow == -1)
                {
                    Log("Id 칼럼을 찾을 수 없어, 시작 행을 구할 수 없습니다.");
                    return false;
                }

                StartCol = 1;

                int EndRow = (worksheet.Cells[StartRow, "A"] as Range).get_End(XlDirection.xlDown).Row;
                int EndCol = (worksheet.Cells[StartRow, "A"] as Range).get_End(XlDirection.xlToRight).Column;

                Log($"시작/마지막 행 : {StartRow}/{EndRow}\n");
                Log($"시작/마지막 열 : {StartCol}/{EndCol}\n");

                int rowCount = EndRow - StartRow + 1;
                int colCount = EndCol;

                Log($"행 갯수 : {rowCount}, ");
                Log($"열 갯수 : {colCount}\n");

                string[] names = new string[colCount];
                string[] types = new string[colCount];
                string[,] datas = new string[rowCount - 2, colCount];
                // Get name / type cells
                Log("데이터 타입과 이름을 가져오는 중...\n");
                for (int i = StartCol; i <= colCount; i++)
                {
                    object name = (worksheet.Cells[StartRow, i] as Range).Value2;
                    if (name == null || name?.ToString() == " ")
                    {
                        names[i - StartCol] = "";
                        Log($"타입 읽기 에러 : 테이블을 확인하세요. 셀 [{StartRow},{i}]");
                    }
                    else
                    {
                        names[i - StartCol] = name.ToString();
                    }

                    Log($"{names[i - StartCol]} : ");

                    object type = (worksheet.Cells[StartRow + 1, i] as Range).Value2;
                    if (type == null || type?.ToString() == " ")
                    {
                        types[i - StartCol] = "";
                        Log($"이름 읽기 에러 : 테이블을 확인하세요. 셀 [{StartRow + 1},{i}]");
                    }
                    else
                    {
                        types[i - StartCol] = type.ToString();
                    }

                    Log($"{types[i - StartCol]}\n");
                }

                // Get data cells
                Log("데이터 읽는 중...\n");
                for (int i = StartCol; i <= colCount; i++)
                {
                    for (int j = StartRow + 2; j <= EndRow; j++)
                    {
                        Log($"{(worksheet.Cells[j, i] as Range).Value2}\n");
                        var read = (worksheet.Cells[j, i] as Range).Value2;
                        if (read == null)
                        {
                            if (types[i - StartCol] == "string")
                            {
                                read = "empty";
                            }
                            else if (types[i - StartCol] == "int" || types[i - StartCol] == "float")
                            {
                                read = 0;
                            }
                        }

                        datas[j - (StartRow + 2), i - (StartCol)] = read?.ToString();
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
                for (int j = 0; j < datas.GetLength(0); j++) // From 0 to row count
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
                Log("JSON 저장 완료.\n");

                Log("C# 클래스 작성중...\n");
                if (Directory.Exists(savePath_Script) == false)
                {
                    Directory.CreateDirectory(savePath_Script);
                }

                targetPath = $"{savePath_Script}\\{fileName[0]}.cs";
                writer = File.CreateText(targetPath);

                writer.WriteLine("// Auto created by table tool. Created by DragonGate Table Loader.");
                writer.WriteLine();
                writer.WriteLine("[System.Serializable]");
                // 클래스 이름
                writer.WriteLine($"public class {fileName[0]} : Data");
                writer.WriteLine("{");
                for (int i = 1; i < datas.GetLength(1); i++) // Id 칼럼은 스킵한다. -> 필요시 Data 상속 제거하고 Id 도 하거나, new 키워드만 붙이거나
                {
                    writer.WriteLine($"\tpublic {types[i]} {names[i]};");
                }

                writer.WriteLine("}");
                writer.WriteLine("[System.Serializable]");
                writer.WriteLine($"public class {fileName[0]}Table : TableBase<{fileName[0]}Data> {{ }}");
                writer.Close();
                Log("C# 클래스 저장 완료.\n");

                // 로컬 마지막 경로 저장
                StreamWriter sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\lastConfig");
                if (sw != null)
                {
                    sw.WriteLine(savePath_Json);
                    sw.WriteLine(savePath_Script);
                    sw.Close();
                }

                return true;
            }
            catch (Exception ex)
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
            if (Directory.Exists(saveJsonPathBox.Text) == true)
            {
                Process.Start("Explorer.exe", saveJsonPathBox.Text);
            }
            else
            {
                Log("잘못된 경로입니다.\n");
            }
        }

        private void OpenClassButton_Click(object sender, EventArgs e)
        {
            if (Directory.Exists(saveClassPathBox.Text) == true)
            {
                Process.Start("Explorer.exe", saveClassPathBox.Text);
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
                saveClassPathBox.Text = fb.SelectedPath;
            }
        }

        // private void pivotApply_Click(object sender, EventArgs e)
        // {
        //     FolderBrowserDialog fb = new FolderBrowserDialog();
        //     if (fb.ShowDialog() == DialogResult.OK)
        //     {
        //         pivotFolderName = fb.SelectedPath;
        //         pivotPathBox.Text = fb.SelectedPath;
        //
        //         if (string.IsNullOrEmpty(pivotPathBox.Text) == true)
        //         {
        //             pivotFolderName = "";
        //             pivotPathBox.Text = pivotFolderName;
        //         }
        //
        //         // 로컬 마지막 경로 저장
        //         StreamWriter sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\lastConfig");
        //         if (sw == null) return;
        //
        //         sw.WriteLine(pivotPathBox.Text);
        //         sw.Close();
        //     }
        //
        //     pathBox_TextChanged(null, null);
        //     fileTextBox_TextChanged(null, null);
        // }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 종료 전에 수행할 작업
            Release();
        }
    }
}