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
            this.excelPathLabel = new System.Windows.Forms.Label();
            this.pathBox = new System.Windows.Forms.TextBox();
            this.findButton = new System.Windows.Forms.Button();
            this.excelFileLabel = new System.Windows.Forms.Label();
            this.GenerateAllButton = new System.Windows.Forms.Button();
            this.fileTextBox = new System.Windows.Forms.TextBox();
            this.GenerateFileButton = new System.Windows.Forms.Button();
            this.fileFindButton = new System.Windows.Forms.Button();
            this.savePathLabel = new System.Windows.Forms.Label();
            this.savePathBox = new System.Windows.Forms.TextBox();
            this.SetSavePathButton = new System.Windows.Forms.Button();
            this.logBox = new System.Windows.Forms.RichTextBox();
            this.logLabel = new System.Windows.Forms.Label();
            this.classFolderFindButton = new System.Windows.Forms.Button();
            this.classPathBox = new System.Windows.Forms.TextBox();
            this.classSavePath = new System.Windows.Forms.Label();
            this.OpenJsonButton = new System.Windows.Forms.Button();
            this.OpenClassButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // excelPathLabel
            // 
            this.excelPathLabel.AutoSize = true;
            this.excelPathLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excelPathLabel.Location = new System.Drawing.Point(14, 20);
            this.excelPathLabel.Name = "excelPathLabel";
            this.excelPathLabel.Size = new System.Drawing.Size(165, 33);
            this.excelPathLabel.TabIndex = 0;
            this.excelPathLabel.Text = "엑셀 폴더 경로";
            // 
            // pathBox
            // 
            this.pathBox.Location = new System.Drawing.Point(21, 56);
            this.pathBox.Multiline = true;
            this.pathBox.Name = "pathBox";
            this.pathBox.Size = new System.Drawing.Size(411, 80);
            this.pathBox.TabIndex = 1;
            this.pathBox.TextChanged += new System.EventHandler(this.pathBox_TextChanged);
            // 
            // findButton
            // 
            this.findButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.findButton.Location = new System.Drawing.Point(438, 56);
            this.findButton.Name = "findButton";
            this.findButton.Size = new System.Drawing.Size(122, 37);
            this.findButton.TabIndex = 2;
            this.findButton.Text = "폴더 선택";
            this.findButton.UseVisualStyleBackColor = true;
            this.findButton.Click += new System.EventHandler(this.FindFolder);
            // 
            // excelFileLabel
            // 
            this.excelFileLabel.AutoSize = true;
            this.excelFileLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.excelFileLabel.Location = new System.Drawing.Point(14, 157);
            this.excelFileLabel.Name = "excelFileLabel";
            this.excelFileLabel.Size = new System.Drawing.Size(165, 33);
            this.excelFileLabel.TabIndex = 3;
            this.excelFileLabel.Text = "엑셀 파일 경로";
            // 
            // GenerateAllButton
            // 
            this.GenerateAllButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GenerateAllButton.Location = new System.Drawing.Point(438, 99);
            this.GenerateAllButton.Name = "GenerateAllButton";
            this.GenerateAllButton.Size = new System.Drawing.Size(122, 37);
            this.GenerateAllButton.TabIndex = 4;
            this.GenerateAllButton.Text = "전체 제작";
            this.GenerateAllButton.UseVisualStyleBackColor = true;
            this.GenerateAllButton.Click += new System.EventHandler(this.GenerateAllButton_Click);
            // 
            // fileTextBox
            // 
            this.fileTextBox.Location = new System.Drawing.Point(21, 193);
            this.fileTextBox.Multiline = true;
            this.fileTextBox.Name = "fileTextBox";
            this.fileTextBox.Size = new System.Drawing.Size(411, 80);
            this.fileTextBox.TabIndex = 5;
            this.fileTextBox.TextChanged += new System.EventHandler(this.fileTextBox_TextChanged);
            // 
            // GenerateFileButton
            // 
            this.GenerateFileButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.GenerateFileButton.Location = new System.Drawing.Point(438, 236);
            this.GenerateFileButton.Name = "GenerateFileButton";
            this.GenerateFileButton.Size = new System.Drawing.Size(122, 37);
            this.GenerateFileButton.TabIndex = 7;
            this.GenerateFileButton.Text = "선택 제작";
            this.GenerateFileButton.UseVisualStyleBackColor = true;
            this.GenerateFileButton.Click += new System.EventHandler(this.GenerateFileButton_Click);
            // 
            // fileFindButton
            // 
            this.fileFindButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fileFindButton.Location = new System.Drawing.Point(438, 193);
            this.fileFindButton.Name = "fileFindButton";
            this.fileFindButton.Size = new System.Drawing.Size(122, 37);
            this.fileFindButton.TabIndex = 6;
            this.fileFindButton.Text = "파일 선택";
            this.fileFindButton.UseVisualStyleBackColor = true;
            this.fileFindButton.Click += new System.EventHandler(this.fileFindButton_Click);
            // 
            // savePathLabel
            // 
            this.savePathLabel.AutoSize = true;
            this.savePathLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.savePathLabel.Location = new System.Drawing.Point(14, 298);
            this.savePathLabel.Name = "savePathLabel";
            this.savePathLabel.Size = new System.Drawing.Size(240, 33);
            this.savePathLabel.TabIndex = 8;
            this.savePathLabel.Text = "JSON 출력 경로 지정";
            // 
            // savePathBox
            // 
            this.savePathBox.Location = new System.Drawing.Point(21, 334);
            this.savePathBox.Multiline = true;
            this.savePathBox.Name = "savePathBox";
            this.savePathBox.Size = new System.Drawing.Size(411, 78);
            this.savePathBox.TabIndex = 9;
            // 
            // SetSavePathButton
            // 
            this.SetSavePathButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SetSavePathButton.Location = new System.Drawing.Point(438, 334);
            this.SetSavePathButton.Name = "SetSavePathButton";
            this.SetSavePathButton.Size = new System.Drawing.Size(122, 36);
            this.SetSavePathButton.TabIndex = 10;
            this.SetSavePathButton.Text = "폴더 선택";
            this.SetSavePathButton.UseVisualStyleBackColor = true;
            this.SetSavePathButton.Click += new System.EventHandler(this.SetSavePathButton_Click);
            // 
            // logBox
            // 
            this.logBox.Font = new System.Drawing.Font("Cascadia Code", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logBox.Location = new System.Drawing.Point(590, 56);
            this.logBox.Name = "logBox";
            this.logBox.ReadOnly = true;
            this.logBox.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.logBox.Size = new System.Drawing.Size(470, 484);
            this.logBox.TabIndex = 11;
            this.logBox.Text = "";
            // 
            // logLabel
            // 
            this.logLabel.AutoSize = true;
            this.logLabel.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.logLabel.Location = new System.Drawing.Point(584, 20);
            this.logLabel.Name = "logLabel";
            this.logLabel.Size = new System.Drawing.Size(110, 33);
            this.logLabel.TabIndex = 12;
            this.logLabel.Text = "실행 로그";
            // 
            // classFolderFindButton
            // 
            this.classFolderFindButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.classFolderFindButton.Location = new System.Drawing.Point(438, 462);
            this.classFolderFindButton.Name = "classFolderFindButton";
            this.classFolderFindButton.Size = new System.Drawing.Size(122, 36);
            this.classFolderFindButton.TabIndex = 15;
            this.classFolderFindButton.Text = "폴더 선택";
            this.classFolderFindButton.UseVisualStyleBackColor = true;
            // 
            // classPathBox
            // 
            this.classPathBox.Location = new System.Drawing.Point(21, 462);
            this.classPathBox.Multiline = true;
            this.classPathBox.Name = "classPathBox";
            this.classPathBox.Size = new System.Drawing.Size(411, 78);
            this.classPathBox.TabIndex = 14;
            // 
            // classSavePath
            // 
            this.classSavePath.AutoSize = true;
            this.classSavePath.Font = new System.Drawing.Font("Cascadia Code", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.classSavePath.Location = new System.Drawing.Point(14, 426);
            this.classSavePath.Name = "classSavePath";
            this.classSavePath.Size = new System.Drawing.Size(300, 33);
            this.classSavePath.TabIndex = 13;
            this.classSavePath.Text = "C# Class 출력 경로 지정";
            // 
            // OpenJsonButton
            // 
            this.OpenJsonButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OpenJsonButton.Location = new System.Drawing.Point(438, 376);
            this.OpenJsonButton.Name = "OpenJsonButton";
            this.OpenJsonButton.Size = new System.Drawing.Size(122, 36);
            this.OpenJsonButton.TabIndex = 16;
            this.OpenJsonButton.Text = "폴더 열기";
            this.OpenJsonButton.UseVisualStyleBackColor = true;
            this.OpenJsonButton.Click += new System.EventHandler(this.OpenJsonButton_Click);
            // 
            // OpenClassButton
            // 
            this.OpenClassButton.Font = new System.Drawing.Font("Cascadia Code", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OpenClassButton.Location = new System.Drawing.Point(438, 504);
            this.OpenClassButton.Name = "OpenClassButton";
            this.OpenClassButton.Size = new System.Drawing.Size(122, 36);
            this.OpenClassButton.TabIndex = 17;
            this.OpenClassButton.Text = "폴더 열기";
            this.OpenClassButton.UseVisualStyleBackColor = true;
            this.OpenClassButton.Click += new System.EventHandler(this.OpenClassButton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.ClientSize = new System.Drawing.Size(1082, 561);
            this.Controls.Add(this.OpenClassButton);
            this.Controls.Add(this.OpenJsonButton);
            this.Controls.Add(this.classFolderFindButton);
            this.Controls.Add(this.classPathBox);
            this.Controls.Add(this.classSavePath);
            this.Controls.Add(this.logLabel);
            this.Controls.Add(this.logBox);
            this.Controls.Add(this.SetSavePathButton);
            this.Controls.Add(this.savePathBox);
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
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Table Generator v1.0 DragonGate";
            this.Load += new System.EventHandler(this.Form1_Load);
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
        private System.Windows.Forms.TextBox savePathBox;
        private System.Windows.Forms.Button SetSavePathButton;
        private System.Windows.Forms.RichTextBox logBox;
        private System.Windows.Forms.Label logLabel;
        private System.Windows.Forms.Button classFolderFindButton;
        private System.Windows.Forms.TextBox classPathBox;
        private System.Windows.Forms.Label classSavePath;
        private System.Windows.Forms.Button OpenJsonButton;
        private System.Windows.Forms.Button OpenClassButton;
    }
}

