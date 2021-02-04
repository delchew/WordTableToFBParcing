
namespace GetInfoFromWordToFireBirdTable
{
    partial class MainForm
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
            this.tableLayoutPanel = new System.Windows.Forms.TableLayoutPanel();
            this.dbFilePathTextBox = new System.Windows.Forms.TextBox();
            this.openDBButton = new System.Windows.Forms.Button();
            this.msWordFilePathTextBox = new System.Windows.Forms.TextBox();
            this.openDocButton = new System.Windows.Forms.Button();
            this.wordTableParseStartButton = new System.Windows.Forms.Button();
            this.cableBrandCheckComboBox = new System.Windows.Forms.ComboBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.tableLayoutPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel
            // 
            this.tableLayoutPanel.ColumnCount = 2;
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 120F));
            this.tableLayoutPanel.Controls.Add(this.dbFilePathTextBox, 0, 0);
            this.tableLayoutPanel.Controls.Add(this.openDBButton, 1, 0);
            this.tableLayoutPanel.Controls.Add(this.msWordFilePathTextBox, 0, 1);
            this.tableLayoutPanel.Controls.Add(this.openDocButton, 1, 1);
            this.tableLayoutPanel.Location = new System.Drawing.Point(12, 12);
            this.tableLayoutPanel.Name = "tableLayoutPanel";
            this.tableLayoutPanel.RowCount = 2;
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel.Size = new System.Drawing.Size(776, 100);
            this.tableLayoutPanel.TabIndex = 0;
            // 
            // dbFilePathTextBox
            // 
            this.dbFilePathTextBox.Location = new System.Drawing.Point(3, 3);
            this.dbFilePathTextBox.Name = "dbFilePathTextBox";
            this.dbFilePathTextBox.Size = new System.Drawing.Size(650, 20);
            this.dbFilePathTextBox.TabIndex = 0;
            // 
            // openDBButton
            // 
            this.openDBButton.Location = new System.Drawing.Point(659, 3);
            this.openDBButton.Name = "openDBButton";
            this.openDBButton.Size = new System.Drawing.Size(114, 23);
            this.openDBButton.TabIndex = 1;
            this.openDBButton.Text = "Open DB";
            this.openDBButton.UseVisualStyleBackColor = true;
            // 
            // msWordFilePathTextBox
            // 
            this.msWordFilePathTextBox.Location = new System.Drawing.Point(3, 53);
            this.msWordFilePathTextBox.Name = "msWordFilePathTextBox";
            this.msWordFilePathTextBox.Size = new System.Drawing.Size(650, 20);
            this.msWordFilePathTextBox.TabIndex = 2;
            // 
            // openDocButton
            // 
            this.openDocButton.Location = new System.Drawing.Point(659, 53);
            this.openDocButton.Name = "openDocButton";
            this.openDocButton.Size = new System.Drawing.Size(114, 23);
            this.openDocButton.TabIndex = 3;
            this.openDocButton.Text = "Open *doc";
            this.openDocButton.UseVisualStyleBackColor = true;
            // 
            // wordTableParseStartButton
            // 
            this.wordTableParseStartButton.Location = new System.Drawing.Point(671, 121);
            this.wordTableParseStartButton.Name = "wordTableParseStartButton";
            this.wordTableParseStartButton.Size = new System.Drawing.Size(114, 23);
            this.wordTableParseStartButton.TabIndex = 1;
            this.wordTableParseStartButton.Text = "Parse!";
            this.wordTableParseStartButton.UseVisualStyleBackColor = true;
            // 
            // cableBrandCheckComboBox
            // 
            this.cableBrandCheckComboBox.FormattingEnabled = true;
            this.cableBrandCheckComboBox.Location = new System.Drawing.Point(12, 123);
            this.cableBrandCheckComboBox.Name = "cableBrandCheckComboBox";
            this.cableBrandCheckComboBox.Size = new System.Drawing.Size(193, 21);
            this.cableBrandCheckComboBox.TabIndex = 2;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(215, 121);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(450, 23);
            this.progressBar.TabIndex = 3;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 153);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.cableBrandCheckComboBox);
            this.Controls.Add(this.wordTableParseStartButton);
            this.Controls.Add(this.tableLayoutPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.Text = "Get Info from Word to Firebird Table";
            this.tableLayoutPanel.ResumeLayout(false);
            this.tableLayoutPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel;
        private System.Windows.Forms.TextBox dbFilePathTextBox;
        private System.Windows.Forms.Button openDBButton;
        private System.Windows.Forms.TextBox msWordFilePathTextBox;
        private System.Windows.Forms.Button openDocButton;
        private System.Windows.Forms.Button wordTableParseStartButton;
        private System.Windows.Forms.ComboBox cableBrandCheckComboBox;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

