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
            pivotPathBox.Text = pivotFolderName;

            if (File.Exists(Directory.GetCurrentDirectory() + "\\lastConfig") == false) return;

            StreamReader sr = new StreamReader(Directory.GetCurrentDirectory() + "\\lastConfig");
            pathBox.Text = sr.ReadLine();
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
                    savePathBox.Text = $"{pathBox.Text}\\Generated";
                    classPathBox.Text = savePathBox.Text;
                }
                else
                {
					savePathBox.Text = $"{pivotFolderName}\\Generated";
					classPathBox.Text = savePathBox.Text;
				}
            }
        }
        private void pathBox_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(pathBox.Text) == true) return;

			if (string.IsNullOrEmpty(pivotFolderName) == true)
			{
				savePathBox.Text = $"{pathBox.Text}\\Generated";
				classPathBox.Text = savePathBox.Text;
			}
			else
			{
				savePathBox.Text = $"{pivotFolderName}\\Generated";
				classPathBox.Text = savePathBox.Text;
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
                    savePathBox.Text = fileTextBox.Text.Substring(0, fileNameStartIndex);
                    savePathBox.Text += "Generated";
                    classPathBox.Text = savePathBox.Text;
				}
				else
				{
					savePathBox.Text = $"{pivotFolderName}\\Generated";
					classPathBox.Text = savePathBox.Text;
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
				savePathBox.Text = fileTextBox.Text.Substring(0, fileNameStartIndex);
				savePathBox.Text += "Generated";
				classPathBox.Text = savePathBox.Text;
			}
			else
			{
				savePathBox.Text = $"{pivotFolderName}\\Generated";
				classPathBox.Text = savePathBox.Text;
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
                if (string.IsNullOrEmpty(savePathBox.Text) == true ||
                    string.IsNullOrEmpty(classPathBox.Text) == true)
                {
                    Log($"저장 경로가 설정되지 않았습니다.\nJson : \"{savePathBox.Text}\"\nClass : \"{classPathBox.Text}\"\n");
                    return;
                }
                if (Directory.Exists(savePathBox.Text) == false)
                {
                    Directory.CreateDirectory(savePathBox.Text);
                }
                if (Directory.Exists(classPathBox.Text) == false)
                {
                    Directory.CreateDirectory(classPathBox.Text);
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
        private int StartRow = 1;
        private bool Make(string argFilePath)
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(argFilePath);
                worksheet = workbook.Worksheets.Item[1] as Worksheet;
                //Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;

                StartRow = (worksheet.Cells[1, "A"] as Range).get_End(XlDirection.xlDown).Row;
                StartCol = 1;

                int EndRow = (worksheet.Cells[StartRow, "A"] as Range).get_End(XlDirection.xlDown).Row;
                int EndCol = (worksheet.Cells[StartRow, "A"] as Range).get_End(XlDirection.xlToRight).Column;

				Log($"시작/마지막 행 : {StartRow}/{EndRow}\n");
				Log($"시작/마지막 열 : {StartCol}/{EndCol}\n");
				//if (range.Cells[StartRow, StartCol] == null || (range.Cells[StartRow, StartCol] as Range).Value2 == null)
				//            {
				//                Log("시작 셀(A3) 이 비어 있습니다.\n");
				//                return false;
				//            }
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
                for (int i = StartCol; i <= EndCol; i++)
                {
                    for (int j = StartRow + 2; j <= EndRow; j++)
                    {
                        Log($"{(worksheet.Cells[j, i] as Range).Value2}\n");
                        datas[j - (StartRow + 2), i - (StartCol)] = (worksheet.Cells[j, i] as Range).Value2.ToString();
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
                writer.WriteLine($"public class {fileName[0]}Data : Data");
                writer.WriteLine("{");
                for (int i = 1; i < datas.GetLength(1); i++) // ID 칼럼은 스킵한다.
                {
                    writer.WriteLine($"\tpublic {types[i]} {names[i]};");
                }
                writer.WriteLine("}");
                writer.WriteLine("[System.Serializable]");
                writer.WriteLine($"public class {fileName[0]}Table : TableBase<{fileName[0]}Data> {{ }}");
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

        private void pivotApply_Click(object sender, EventArgs e)
        {
			FolderBrowserDialog fb = new FolderBrowserDialog();
			if (fb.ShowDialog() == DialogResult.OK)
			{
                pivotFolderName = fb.SelectedPath;
				pivotPathBox.Text = fb.SelectedPath;

				if (string.IsNullOrEmpty(pivotPathBox.Text) == true)
				{
					pivotFolderName = "";
					pivotPathBox.Text = pivotFolderName;
				}
				// 로컬 마지막 경로 저장
				StreamWriter sw = new StreamWriter(Directory.GetCurrentDirectory() + "\\lastConfig");
				if (sw == null) return;
                
				sw.WriteLine(pivotPathBox.Text);
				sw.Close();
			}
            pathBox_TextChanged(null, null);
            fileTextBox_TextChanged(null, null);
        }
    }
}
