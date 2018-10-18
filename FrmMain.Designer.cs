namespace _3VJ_MV
{
    partial class FrmMain
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
            this.btnForm1 = new System.Windows.Forms.Button();
            this.btnFormInvoke = new System.Windows.Forms.Button();
            this.txt1 = new System.Windows.Forms.TextBox();
            this.txt2 = new System.Windows.Forms.TextBox();
            this.btn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnForm1
            // 
            this.btnForm1.Location = new System.Drawing.Point(75, 53);
            this.btnForm1.Name = "btnForm1";
            this.btnForm1.Size = new System.Drawing.Size(114, 23);
            this.btnForm1.TabIndex = 0;
            this.btnForm1.Text = "打开Form1";
            this.btnForm1.UseVisualStyleBackColor = true;
            this.btnForm1.Click += new System.EventHandler(this.btnForm1_Click);
            // 
            // btnFormInvoke
            // 
            this.btnFormInvoke.Location = new System.Drawing.Point(75, 140);
            this.btnFormInvoke.Name = "btnFormInvoke";
            this.btnFormInvoke.Size = new System.Drawing.Size(114, 23);
            this.btnFormInvoke.TabIndex = 0;
            this.btnFormInvoke.Text = "打开FormInvoke";
            this.btnFormInvoke.UseVisualStyleBackColor = true;
            this.btnFormInvoke.Click += new System.EventHandler(this.btnFormInvoke_Click);
            // 
            // txt1
            // 
            this.txt1.Location = new System.Drawing.Point(75, 83);
            this.txt1.Name = "txt1";
            this.txt1.Size = new System.Drawing.Size(378, 21);
            this.txt1.TabIndex = 1;
            // 
            // txt2
            // 
            this.txt2.Location = new System.Drawing.Point(75, 169);
            this.txt2.Name = "txt2";
            this.txt2.Size = new System.Drawing.Size(378, 21);
            this.txt2.TabIndex = 1;
            // 
            // btn
            // 
            this.btn.Location = new System.Drawing.Point(222, 218);
            this.btn.Name = "btn";
            this.btn.Size = new System.Drawing.Size(75, 23);
            this.btn.TabIndex = 2;
            this.btn.Text = "对比CSV";
            this.btn.UseVisualStyleBackColor = true;
            this.btn.Click += new System.EventHandler(this.btn_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(554, 364);
            this.Controls.Add(this.btn);
            this.Controls.Add(this.txt2);
            this.Controls.Add(this.txt1);
            this.Controls.Add(this.btnFormInvoke);
            this.Controls.Add(this.btnForm1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FrmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmMain";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnForm1;
        private System.Windows.Forms.Button btnFormInvoke;
        private System.Windows.Forms.TextBox txt1;
        private System.Windows.Forms.TextBox txt2;
        private System.Windows.Forms.Button btn;
    }
}