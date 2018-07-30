namespace WindowsFormsApp1
{
    partial class Form1
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
            this.btnMerge = new System.Windows.Forms.Button();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnDocXTest = new System.Windows.Forms.Button();
            this.btnDocXLoadFile = new System.Windows.Forms.Button();
            this.btnTemplateEditorTest = new System.Windows.Forms.Button();
            this.btnSchrijfOfferteRegel = new System.Windows.Forms.Button();
            this.btnSearchAndShow = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnMerge
            // 
            this.btnMerge.Location = new System.Drawing.Point(171, 77);
            this.btnMerge.Name = "btnMerge";
            this.btnMerge.Size = new System.Drawing.Size(75, 23);
            this.btnMerge.TabIndex = 0;
            this.btnMerge.Text = "Merge";
            this.btnMerge.UseVisualStyleBackColor = true;
            this.btnMerge.Click += new System.EventHandler(this.btnMerge_Click);
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(335, 94);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(75, 23);
            this.btnOpenFile.TabIndex = 1;
            this.btnOpenFile.Text = "Open file";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // btnDocXTest
            // 
            this.btnDocXTest.Location = new System.Drawing.Point(152, 202);
            this.btnDocXTest.Name = "btnDocXTest";
            this.btnDocXTest.Size = new System.Drawing.Size(75, 23);
            this.btnDocXTest.TabIndex = 2;
            this.btnDocXTest.Text = "DocX Test";
            this.btnDocXTest.UseVisualStyleBackColor = true;
            this.btnDocXTest.Click += new System.EventHandler(this.btnDocXTest_Click);
            // 
            // btnDocXLoadFile
            // 
            this.btnDocXLoadFile.Location = new System.Drawing.Point(547, 169);
            this.btnDocXLoadFile.Name = "btnDocXLoadFile";
            this.btnDocXLoadFile.Size = new System.Drawing.Size(75, 23);
            this.btnDocXLoadFile.TabIndex = 3;
            this.btnDocXLoadFile.Text = "DocX Load File";
            this.btnDocXLoadFile.UseVisualStyleBackColor = true;
            this.btnDocXLoadFile.Click += new System.EventHandler(this.btnDocXLoadFile_Click);
            // 
            // btnTemplateEditorTest
            // 
            this.btnTemplateEditorTest.Location = new System.Drawing.Point(547, 256);
            this.btnTemplateEditorTest.Name = "btnTemplateEditorTest";
            this.btnTemplateEditorTest.Size = new System.Drawing.Size(138, 23);
            this.btnTemplateEditorTest.TabIndex = 4;
            this.btnTemplateEditorTest.Text = "Template editor test";
            this.btnTemplateEditorTest.UseVisualStyleBackColor = true;
            this.btnTemplateEditorTest.Click += new System.EventHandler(this.btnTemplateEditorTest_Click);
            // 
            // btnSchrijfOfferteRegel
            // 
            this.btnSchrijfOfferteRegel.Location = new System.Drawing.Point(383, 323);
            this.btnSchrijfOfferteRegel.Name = "btnSchrijfOfferteRegel";
            this.btnSchrijfOfferteRegel.Size = new System.Drawing.Size(134, 23);
            this.btnSchrijfOfferteRegel.TabIndex = 5;
            this.btnSchrijfOfferteRegel.Text = "Schrijf offerte regel";
            this.btnSchrijfOfferteRegel.UseVisualStyleBackColor = true;
            this.btnSchrijfOfferteRegel.Click += new System.EventHandler(this.btnSchrijfOfferteRegel_Click);
            // 
            // btnSearchAndShow
            // 
            this.btnSearchAndShow.Location = new System.Drawing.Point(66, 231);
            this.btnSearchAndShow.Name = "btnSearchAndShow";
            this.btnSearchAndShow.Size = new System.Drawing.Size(161, 23);
            this.btnSearchAndShow.TabIndex = 6;
            this.btnSearchAndShow.Text = "Zoek repeat blok en laat zien";
            this.btnSearchAndShow.UseVisualStyleBackColor = true;
            this.btnSearchAndShow.Click += new System.EventHandler(this.btnSearchAndShow_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnSearchAndShow);
            this.Controls.Add(this.btnSchrijfOfferteRegel);
            this.Controls.Add(this.btnTemplateEditorTest);
            this.Controls.Add(this.btnDocXLoadFile);
            this.Controls.Add(this.btnDocXTest);
            this.Controls.Add(this.btnOpenFile);
            this.Controls.Add(this.btnMerge);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnMerge;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnDocXTest;
        private System.Windows.Forms.Button btnDocXLoadFile;
        private System.Windows.Forms.Button btnTemplateEditorTest;
        private System.Windows.Forms.Button btnSchrijfOfferteRegel;
        private System.Windows.Forms.Button btnSearchAndShow;
    }
}

