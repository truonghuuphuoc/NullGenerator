namespace NULL_is_my_son
{
    partial class MainAppForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainAppForm));
            this.btnInputFile = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.edtFileInput = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.edtDatabasePath = new System.Windows.Forms.TextBox();
            this.btnBowseDatabase = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.cbbCommandType = new System.Windows.Forms.ComboBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.edtOutFilePath = new System.Windows.Forms.TextBox();
            this.btnBrowseOutFile = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.btnGenerateFile = new System.Windows.Forms.Button();
            this.btnClearLog = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnInputFile
            // 
            resources.ApplyResources(this.btnInputFile, "btnInputFile");
            this.btnInputFile.Name = "btnInputFile";
            this.btnInputFile.UseVisualStyleBackColor = true;
            this.btnInputFile.Click += new System.EventHandler(this.btnInputFile_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.edtFileInput);
            this.groupBox1.Controls.Add(this.btnInputFile);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // edtFileInput
            // 
            resources.ApplyResources(this.edtFileInput, "edtFileInput");
            this.edtFileInput.Name = "edtFileInput";
            this.edtFileInput.ReadOnly = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.edtDatabasePath);
            this.groupBox2.Controls.Add(this.btnBowseDatabase);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // edtDatabasePath
            // 
            resources.ApplyResources(this.edtDatabasePath, "edtDatabasePath");
            this.edtDatabasePath.Name = "edtDatabasePath";
            this.edtDatabasePath.ReadOnly = true;
            // 
            // btnBowseDatabase
            // 
            resources.ApplyResources(this.btnBowseDatabase, "btnBowseDatabase");
            this.btnBowseDatabase.Name = "btnBowseDatabase";
            this.btnBowseDatabase.UseVisualStyleBackColor = true;
            this.btnBowseDatabase.Click += new System.EventHandler(this.btnBowseDatabase_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.cbbCommandType);
            resources.ApplyResources(this.groupBox3, "groupBox3");
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.TabStop = false;
            // 
            // cbbCommandType
            // 
            this.cbbCommandType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbbCommandType.FormattingEnabled = true;
            this.cbbCommandType.Items.AddRange(new object[] {
            resources.GetString("cbbCommandType.Items"),
            resources.GetString("cbbCommandType.Items1")});
            resources.ApplyResources(this.cbbCommandType, "cbbCommandType");
            this.cbbCommandType.Name = "cbbCommandType";
            this.cbbCommandType.SelectedIndexChanged += new System.EventHandler(this.cbbCommandType_SelectedIndexChanged);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.edtOutFilePath);
            this.groupBox4.Controls.Add(this.btnBrowseOutFile);
            resources.ApplyResources(this.groupBox4, "groupBox4");
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.TabStop = false;
            // 
            // edtOutFilePath
            // 
            resources.ApplyResources(this.edtOutFilePath, "edtOutFilePath");
            this.edtOutFilePath.Name = "edtOutFilePath";
            this.edtOutFilePath.ReadOnly = true;
            // 
            // btnBrowseOutFile
            // 
            resources.ApplyResources(this.btnBrowseOutFile, "btnBrowseOutFile");
            this.btnBrowseOutFile.Name = "btnBrowseOutFile";
            this.btnBrowseOutFile.UseVisualStyleBackColor = true;
            this.btnBrowseOutFile.Click += new System.EventHandler(this.btnBrowseOutFile_Click);
            // 
            // txtLog
            // 
            resources.ApplyResources(this.txtLog, "txtLog");
            this.txtLog.Name = "txtLog";
            this.txtLog.ReadOnly = true;
            // 
            // btnGenerateFile
            // 
            resources.ApplyResources(this.btnGenerateFile, "btnGenerateFile");
            this.btnGenerateFile.Name = "btnGenerateFile";
            this.btnGenerateFile.UseVisualStyleBackColor = true;
            this.btnGenerateFile.Click += new System.EventHandler(this.btnGenerateFile_Click);
            // 
            // btnClearLog
            // 
            resources.ApplyResources(this.btnClearLog, "btnClearLog");
            this.btnClearLog.Name = "btnClearLog";
            this.btnClearLog.UseVisualStyleBackColor = true;
            this.btnClearLog.Click += new System.EventHandler(this.btnClearLog_Click);
            // 
            // MainAppForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.MenuBar;
            this.Controls.Add(this.btnClearLog);
            this.Controls.Add(this.btnGenerateFile);
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.Name = "MainAppForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnInputFile;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox edtFileInput;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox edtDatabasePath;
        private System.Windows.Forms.Button btnBowseDatabase;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.TextBox edtOutFilePath;
        private System.Windows.Forms.Button btnBrowseOutFile;
        private System.Windows.Forms.ComboBox cbbCommandType;

        private int intCommandType = 0;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Button btnGenerateFile;
        private System.Windows.Forms.Button btnClearLog;
    }
}

