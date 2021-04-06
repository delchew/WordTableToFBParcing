
namespace WordTableToFBParsing
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
            this.tableLayoutPanelTop = new System.Windows.Forms.TableLayoutPanel();
            this.openDocButton = new System.Windows.Forms.Button();
            this.msWordFilePathTextBox = new System.Windows.Forms.TextBox();
            this.dbConnectionCheckComboBox = new System.Windows.Forms.ComboBox();
            this.tableLayoutPanelBottom = new System.Windows.Forms.TableLayoutPanel();
            this.cableBrandCheckComboBox = new System.Windows.Forms.ComboBox();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.wordTableParseStartButton = new System.Windows.Forms.Button();
            this.tableLayoutPanelTop.SuspendLayout();
            this.tableLayoutPanelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanelTop
            // 
            this.tableLayoutPanelTop.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanelTop.ColumnCount = 2;
            this.tableLayoutPanelTop.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 83.35484F));
            this.tableLayoutPanelTop.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 16.64516F));
            this.tableLayoutPanelTop.Controls.Add(this.openDocButton, 1, 0);
            this.tableLayoutPanelTop.Controls.Add(this.msWordFilePathTextBox, 0, 0);
            this.tableLayoutPanelTop.Controls.Add(this.dbConnectionCheckComboBox, 0, 1);
            this.tableLayoutPanelTop.Location = new System.Drawing.Point(13, 13);
            this.tableLayoutPanelTop.Name = "tableLayoutPanelTop";
            this.tableLayoutPanelTop.RowCount = 2;
            this.tableLayoutPanelTop.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 52.17391F));
            this.tableLayoutPanelTop.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 47.82609F));
            this.tableLayoutPanelTop.Size = new System.Drawing.Size(759, 69);
            this.tableLayoutPanelTop.TabIndex = 0;
            // 
            // openDocButton
            // 
            this.openDocButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.openDocButton.Location = new System.Drawing.Point(635, 3);
            this.openDocButton.Name = "openDocButton";
            this.openDocButton.Size = new System.Drawing.Size(121, 29);
            this.openDocButton.TabIndex = 1;
            this.openDocButton.Text = "Open document";
            this.openDocButton.UseVisualStyleBackColor = true;
            // 
            // msWordFilePathTextBox
            // 
            this.msWordFilePathTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.msWordFilePathTextBox.Location = new System.Drawing.Point(3, 7);
            this.msWordFilePathTextBox.Name = "msWordFilePathTextBox";
            this.msWordFilePathTextBox.Size = new System.Drawing.Size(626, 20);
            this.msWordFilePathTextBox.TabIndex = 0;
            // 
            // dbConnectionCheckComboBox
            // 
            this.dbConnectionCheckComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.dbConnectionCheckComboBox.FormattingEnabled = true;
            this.dbConnectionCheckComboBox.Location = new System.Drawing.Point(3, 41);
            this.dbConnectionCheckComboBox.Name = "dbConnectionCheckComboBox";
            this.dbConnectionCheckComboBox.Size = new System.Drawing.Size(626, 21);
            this.dbConnectionCheckComboBox.TabIndex = 2;
            // 
            // tableLayoutPanelBottom
            // 
            this.tableLayoutPanelBottom.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanelBottom.ColumnCount = 3;
            this.tableLayoutPanelBottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.90333F));
            this.tableLayoutPanelBottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 69.09667F));
            this.tableLayoutPanelBottom.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 127F));
            this.tableLayoutPanelBottom.Controls.Add(this.cableBrandCheckComboBox, 0, 0);
            this.tableLayoutPanelBottom.Controls.Add(this.progressBar, 1, 0);
            this.tableLayoutPanelBottom.Controls.Add(this.wordTableParseStartButton, 2, 0);
            this.tableLayoutPanelBottom.Location = new System.Drawing.Point(13, 114);
            this.tableLayoutPanelBottom.Name = "tableLayoutPanelBottom";
            this.tableLayoutPanelBottom.RowCount = 1;
            this.tableLayoutPanelBottom.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanelBottom.Size = new System.Drawing.Size(759, 35);
            this.tableLayoutPanelBottom.TabIndex = 1;
            // 
            // cableBrandCheckComboBox
            // 
            this.cableBrandCheckComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.cableBrandCheckComboBox.FormattingEnabled = true;
            this.cableBrandCheckComboBox.Location = new System.Drawing.Point(3, 7);
            this.cableBrandCheckComboBox.Name = "cableBrandCheckComboBox";
            this.cableBrandCheckComboBox.Size = new System.Drawing.Size(189, 21);
            this.cableBrandCheckComboBox.TabIndex = 0;
            // 
            // progressBar
            // 
            this.progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar.Location = new System.Drawing.Point(198, 6);
            this.progressBar.Maximum = 1;
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(430, 23);
            this.progressBar.TabIndex = 1;
            // 
            // wordTableParseStartButton
            // 
            this.wordTableParseStartButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.wordTableParseStartButton.Location = new System.Drawing.Point(634, 3);
            this.wordTableParseStartButton.Name = "wordTableParseStartButton";
            this.wordTableParseStartButton.Size = new System.Drawing.Size(122, 29);
            this.wordTableParseStartButton.TabIndex = 2;
            this.wordTableParseStartButton.Text = "Parse!";
            this.wordTableParseStartButton.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 161);
            this.Controls.Add(this.tableLayoutPanelBottom);
            this.Controls.Add(this.tableLayoutPanelTop);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Parse MS Word table to Firebird database";
            this.tableLayoutPanelTop.ResumeLayout(false);
            this.tableLayoutPanelTop.PerformLayout();
            this.tableLayoutPanelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelTop;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelBottom;
        private System.Windows.Forms.TextBox msWordFilePathTextBox;
        private System.Windows.Forms.Button openDocButton;
        private System.Windows.Forms.ComboBox dbConnectionCheckComboBox;
        private System.Windows.Forms.ComboBox cableBrandCheckComboBox;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button wordTableParseStartButton;
    }
}

