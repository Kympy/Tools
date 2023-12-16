using System.Windows.Forms;

namespace TableLoader
{
    partial class Form1
    {
        /// <summary>
        /// 필수 디자이너 변수입니다.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 사용 중인 모든 리소스를 정리합니다.
        /// </summary>
        /// <param name="disposing">관리되는 리소스를 삭제해야 하면 true이고, 그렇지 않으면 false입니다.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 디자이너에서 생성한 코드

        /// <summary>
        /// 디자이너 지원에 필요한 메서드입니다. 
        /// 이 메서드의 내용을 코드 편집기로 수정하지 마세요.
        /// </summary>
        private void InitializeComponent()
        {
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
			this.excelPathLabel = new System.Windows.Forms.Label();
			this.pathBox = new System.Windows.Forms.TextBox();
			this.findButton = new System.Windows.Forms.Button();
			this.excelFileLabel = new System.Windows.Forms.Label();
			this.GenerateAllButton = new System.Windows.Forms.Button();
			this.fileTextBox = new System.Windows.Forms.TextBox();
			this.GenerateFileButton = new System.Windows.Forms.Button();
			this.fileFindButton = new System.Windows.Forms.Button();
			this.savePathLabel = new System.Windows.Forms.Label();
			this.saveJsonPathBox = new System.Windows.Forms.TextBox();
			this.SetSavePathButton = new System.Windows.Forms.Button();
			this.logBox = new System.Windows.Forms.RichTextBox();
			this.logLabel = new System.Windows.Forms.Label();
			this.classFolderFindButton = new System.Windows.Forms.Button();
			this.saveClassPathBox = new System.Windows.Forms.TextBox();
			this.classSavePath = new System.Windows.Forms.Label();
			this.OpenJsonButton = new System.Windows.Forms.Button();
			this.OpenClassButton = new System.Windows.Forms.Button();
			//this.pivotFolderLabel = new System.Windows.Forms.Label();
			//this.pivotPathBox = new System.Windows.Forms.TextBox();
			//this.pivotApply = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// excelPathLabel
			// 
			this.excelPathLabel.AutoSize = true;
			this.excelPathLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.excelPathLabel.Location = new System.Drawing.Point(12, 16);
			this.excelPathLabel.Name = "excelPathLabel";
			this.excelPathLabel.Size = new System.Drawing.Size(126, 27);
			this.excelPathLabel.TabIndex = 0;
			this.excelPathLabel.Text = "엑셀 폴더 경로";
			// 
			// pathBox
			// 
			this.pathBox.BackColor = System.Drawing.Color.MistyRose;
			this.pathBox.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.pathBox.Location = new System.Drawing.Point(18, 45);
			this.pathBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.pathBox.Multiline = true;
			this.pathBox.Name = "pathBox";
			this.pathBox.Size = new System.Drawing.Size(360, 65);
			this.pathBox.TabIndex = 1;
			this.pathBox.TextChanged += new System.EventHandler(this.pathBox_TextChanged);
			// 
			// findButton
			// 
			this.findButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.findButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.findButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.findButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.findButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.findButton.Location = new System.Drawing.Point(383, 45);
			this.findButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.findButton.Name = "findButton";
			this.findButton.Size = new System.Drawing.Size(107, 30);
			this.findButton.TabIndex = 2;
			this.findButton.Text = "폴더 선택";
			this.findButton.UseVisualStyleBackColor = true;
			this.findButton.Click += new System.EventHandler(this.FindFolder);
			// 
			// excelFileLabel
			// 
			this.excelFileLabel.AutoSize = true;
			this.excelFileLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.excelFileLabel.Location = new System.Drawing.Point(12, 126);
			this.excelFileLabel.Name = "excelFileLabel";
			this.excelFileLabel.Size = new System.Drawing.Size(126, 27);
			this.excelFileLabel.TabIndex = 3;
			this.excelFileLabel.Text = "엑셀 파일 경로";
			// 
			// GenerateAllButton
			// 
			this.GenerateAllButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.GenerateAllButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.GenerateAllButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.GenerateAllButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.GenerateAllButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.GenerateAllButton.Location = new System.Drawing.Point(383, 79);
			this.GenerateAllButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.GenerateAllButton.Name = "GenerateAllButton";
			this.GenerateAllButton.Size = new System.Drawing.Size(107, 30);
			this.GenerateAllButton.TabIndex = 4;
			this.GenerateAllButton.Text = "전체 제작";
			this.GenerateAllButton.UseVisualStyleBackColor = true;
			this.GenerateAllButton.Click += new System.EventHandler(this.GenerateAllButton_Click);
			// 
			// fileTextBox
			// 
			this.fileTextBox.BackColor = System.Drawing.Color.PeachPuff;
			this.fileTextBox.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.fileTextBox.Location = new System.Drawing.Point(18, 154);
			this.fileTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.fileTextBox.Multiline = true;
			this.fileTextBox.Name = "fileTextBox";
			this.fileTextBox.Size = new System.Drawing.Size(360, 65);
			this.fileTextBox.TabIndex = 5;
			this.fileTextBox.TextChanged += new System.EventHandler(this.fileTextBox_TextChanged);
			// 
			// GenerateFileButton
			// 
			this.GenerateFileButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.GenerateFileButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.GenerateFileButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.GenerateFileButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.GenerateFileButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.GenerateFileButton.Location = new System.Drawing.Point(383, 189);
			this.GenerateFileButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.GenerateFileButton.Name = "GenerateFileButton";
			this.GenerateFileButton.Size = new System.Drawing.Size(107, 30);
			this.GenerateFileButton.TabIndex = 7;
			this.GenerateFileButton.Text = "선택 제작";
			this.GenerateFileButton.UseVisualStyleBackColor = true;
			this.GenerateFileButton.Click += new System.EventHandler(this.GenerateFileButton_Click);
			// 
			// fileFindButton
			// 
			this.fileFindButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.fileFindButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.fileFindButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.fileFindButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.fileFindButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.fileFindButton.Location = new System.Drawing.Point(383, 154);
			this.fileFindButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.fileFindButton.Name = "fileFindButton";
			this.fileFindButton.Size = new System.Drawing.Size(107, 30);
			this.fileFindButton.TabIndex = 6;
			this.fileFindButton.Text = "파일 선택";
			this.fileFindButton.UseVisualStyleBackColor = true;
			this.fileFindButton.Click += new System.EventHandler(this.fileFindButton_Click);
			// 
			// savePathLabel
			// 
			this.savePathLabel.AutoSize = true;
			this.savePathLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.savePathLabel.Location = new System.Drawing.Point(12, 238);
			this.savePathLabel.Name = "savePathLabel";
			this.savePathLabel.Size = new System.Drawing.Size(186, 27);
			this.savePathLabel.TabIndex = 8;
			this.savePathLabel.Text = "JSON 출력 경로 지정";
			// 
			// savePathBox
			// 
			this.saveJsonPathBox.BackColor = System.Drawing.Color.OldLace;
			this.saveJsonPathBox.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.saveJsonPathBox.Location = new System.Drawing.Point(18, 267);
			this.saveJsonPathBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.saveJsonPathBox.Multiline = true;
			this.saveJsonPathBox.Name = "saveJsonPathBox";
			this.saveJsonPathBox.Size = new System.Drawing.Size(360, 63);
			this.saveJsonPathBox.TabIndex = 9;
			// 
			// SetSavePathButton
			// 
			this.SetSavePathButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.SetSavePathButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.SetSavePathButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.SetSavePathButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.SetSavePathButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.SetSavePathButton.Location = new System.Drawing.Point(383, 267);
			this.SetSavePathButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.SetSavePathButton.Name = "SetSavePathButton";
			this.SetSavePathButton.Size = new System.Drawing.Size(107, 29);
			this.SetSavePathButton.TabIndex = 10;
			this.SetSavePathButton.Text = "폴더 선택";
			this.SetSavePathButton.UseVisualStyleBackColor = true;
			this.SetSavePathButton.Click += new System.EventHandler(this.SetSavePathButton_Click);
			// 
			// logBox
			// 
			this.logBox.BackColor = System.Drawing.Color.LavenderBlush;
			this.logBox.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.logBox.Location = new System.Drawing.Point(516, 45);
			this.logBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.logBox.Name = "logBox";
			this.logBox.ReadOnly = true;
			this.logBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
			this.logBox.Size = new System.Drawing.Size(412, 354);
			this.logBox.TabIndex = 11;
			this.logBox.Text = "";
			// 
			// logLabel
			// 
			this.logLabel.AutoSize = true;
			this.logLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.logLabel.Location = new System.Drawing.Point(511, 16);
			this.logLabel.Name = "logLabel";
			this.logLabel.Size = new System.Drawing.Size(84, 27);
			this.logLabel.TabIndex = 12;
			this.logLabel.Text = "실행 로그";
			// 
			// classFolderFindButton
			// 
			this.classFolderFindButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.classFolderFindButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.classFolderFindButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.classFolderFindButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.classFolderFindButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.classFolderFindButton.Location = new System.Drawing.Point(383, 370);
			this.classFolderFindButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.classFolderFindButton.Name = "classFolderFindButton";
			this.classFolderFindButton.Size = new System.Drawing.Size(107, 29);
			this.classFolderFindButton.TabIndex = 15;
			this.classFolderFindButton.Text = "폴더 선택";
			this.classFolderFindButton.UseVisualStyleBackColor = true;
			this.classFolderFindButton.Click += new System.EventHandler(this.classFolderFindButton_Click);
			// 
			// classPathBox
			// 
			this.saveClassPathBox.BackColor = System.Drawing.Color.LemonChiffon;
			this.saveClassPathBox.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.saveClassPathBox.Location = new System.Drawing.Point(18, 370);
			this.saveClassPathBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.saveClassPathBox.Multiline = true;
			this.saveClassPathBox.Name = "saveClassPathBox";
			this.saveClassPathBox.Size = new System.Drawing.Size(360, 63);
			this.saveClassPathBox.TabIndex = 14;
			// 
			// classSavePath
			// 
			this.classSavePath.AutoSize = true;
			this.classSavePath.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.classSavePath.Location = new System.Drawing.Point(12, 341);
			this.classSavePath.Name = "classSavePath";
			this.classSavePath.Size = new System.Drawing.Size(234, 27);
			this.classSavePath.TabIndex = 13;
			this.classSavePath.Text = "C# Class 출력 경로 지정";
			// 
			// OpenJsonButton
			// 
			this.OpenJsonButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.OpenJsonButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.OpenJsonButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.OpenJsonButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.OpenJsonButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.OpenJsonButton.Location = new System.Drawing.Point(383, 301);
			this.OpenJsonButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.OpenJsonButton.Name = "OpenJsonButton";
			this.OpenJsonButton.Size = new System.Drawing.Size(107, 29);
			this.OpenJsonButton.TabIndex = 16;
			this.OpenJsonButton.Text = "폴더 열기";
			this.OpenJsonButton.UseVisualStyleBackColor = true;
			this.OpenJsonButton.Click += new System.EventHandler(this.OpenJsonButton_Click);
			// 
			// OpenClassButton
			// 
			this.OpenClassButton.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			this.OpenClassButton.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			this.OpenClassButton.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			this.OpenClassButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			this.OpenClassButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.OpenClassButton.Location = new System.Drawing.Point(383, 403);
			this.OpenClassButton.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.OpenClassButton.Name = "OpenClassButton";
			this.OpenClassButton.Size = new System.Drawing.Size(107, 29);
			this.OpenClassButton.TabIndex = 17;
			this.OpenClassButton.Text = "폴더 열기";
			this.OpenClassButton.UseVisualStyleBackColor = true;
			this.OpenClassButton.Click += new System.EventHandler(this.OpenClassButton_Click);
			// 
			// pivotFolderLabel
			// 
			// this.pivotFolderLabel.AutoSize = true;
			// this.pivotFolderLabel.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			// this.pivotFolderLabel.Location = new System.Drawing.Point(501, 409);
			// this.pivotFolderLabel.Name = "pivotFolderLabel";
			// this.pivotFolderLabel.Size = new System.Drawing.Size(105, 16);
			// this.pivotFolderLabel.TabIndex = 18;
			// this.pivotFolderLabel.Text = "커스텀 저장 폴더";
			// 
			// pivotPathBox
			// 
			// this.pivotPathBox.BackColor = System.Drawing.Color.AliceBlue;
			// this.pivotPathBox.Font = new System.Drawing.Font("Cascadia Code", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			// this.pivotPathBox.Location = new System.Drawing.Point(612, 406);
			// this.pivotPathBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			// this.pivotPathBox.Name = "pivotPathBox";
			// this.pivotPathBox.ReadOnly = true;
			// this.pivotPathBox.Size = new System.Drawing.Size(236, 21);
			// this.pivotPathBox.TabIndex = 19;
			// this.pivotPathBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			// 
			// pivotApply
			// 
			// this.pivotApply.FlatAppearance.BorderColor = System.Drawing.Color.OliveDrab;
			// this.pivotApply.FlatAppearance.MouseDownBackColor = System.Drawing.Color.PaleGreen;
			// this.pivotApply.FlatAppearance.MouseOverBackColor = System.Drawing.Color.SpringGreen;
			// this.pivotApply.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
			// this.pivotApply.Font = new System.Drawing.Font("Cascadia Code", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			// this.pivotApply.Location = new System.Drawing.Point(854, 403);
			// this.pivotApply.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			// this.pivotApply.Name = "pivotApply";
			// this.pivotApply.Size = new System.Drawing.Size(73, 26);
			// this.pivotApply.TabIndex = 20;
			// this.pivotApply.Text = "찾기";
			// this.pivotApply.UseVisualStyleBackColor = true;
			// this.pivotApply.Click += new System.EventHandler(this.pivotApply_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.Color.White;
			this.ClientSize = new System.Drawing.Size(947, 449);
			// this.Controls.Add(this.pivotApply);
			// this.Controls.Add(this.pivotPathBox);
			// this.Controls.Add(this.pivotFolderLabel);
			this.Controls.Add(this.OpenClassButton);
			this.Controls.Add(this.OpenJsonButton);
			this.Controls.Add(this.classFolderFindButton);
			this.Controls.Add(this.saveClassPathBox);
			this.Controls.Add(this.classSavePath);
			this.Controls.Add(this.logLabel);
			this.Controls.Add(this.logBox);
			this.Controls.Add(this.SetSavePathButton);
			this.Controls.Add(this.saveJsonPathBox);
			this.Controls.Add(this.savePathLabel);
			this.Controls.Add(this.GenerateFileButton);
			this.Controls.Add(this.fileFindButton);
			this.Controls.Add(this.fileTextBox);
			this.Controls.Add(this.GenerateAllButton);
			this.Controls.Add(this.excelFileLabel);
			this.Controls.Add(this.findButton);
			this.Controls.Add(this.pathBox);
			this.Controls.Add(this.excelPathLabel);
			this.Cursor = System.Windows.Forms.Cursors.Default;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.MaximizeBox = false;
			this.Name = "Form1";
			this.Text = "Table Generator v1.0 DragonGate";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.FormClosing += Form1_FormClosing;
			this.ResumeLayout(false);
			this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label excelPathLabel;
        private System.Windows.Forms.TextBox pathBox;
        private System.Windows.Forms.Button findButton;
        private System.Windows.Forms.Label excelFileLabel;
        private System.Windows.Forms.Button GenerateAllButton;
        private System.Windows.Forms.TextBox fileTextBox;
        private System.Windows.Forms.Button GenerateFileButton;
        private System.Windows.Forms.Button fileFindButton;
        private System.Windows.Forms.Label savePathLabel;
        private System.Windows.Forms.TextBox saveJsonPathBox;
        private System.Windows.Forms.Button SetSavePathButton;
        private System.Windows.Forms.RichTextBox logBox;
        private System.Windows.Forms.Label logLabel;
        private System.Windows.Forms.Button classFolderFindButton;
        private System.Windows.Forms.TextBox saveClassPathBox;
        private System.Windows.Forms.Label classSavePath;
        private System.Windows.Forms.Button OpenJsonButton;
        private System.Windows.Forms.Button OpenClassButton;
        //private System.Windows.Forms.Label pivotFolderLabel;
        //private System.Windows.Forms.TextBox pivotPathBox;
        //private System.Windows.Forms.Button pivotApply;
    }
}

