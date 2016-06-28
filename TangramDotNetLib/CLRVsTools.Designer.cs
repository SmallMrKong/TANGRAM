namespace TangramDotNetLib
{
    partial class CLRVsTools
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.ShowWinForm = new System.Windows.Forms.Button();
            this.ShowWinMDIForm = new System.Windows.Forms.Button();
            this.ShowAsCtrl = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ShowWinForm
            // 
            this.ShowWinForm.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ShowWinForm.Location = new System.Drawing.Point(22, 15);
            this.ShowWinForm.Name = "ShowWinForm";
            this.ShowWinForm.Size = new System.Drawing.Size(286, 29);
            this.ShowWinForm.TabIndex = 0;
            this.ShowWinForm.Text = "ShowWinForm";
            this.ShowWinForm.UseVisualStyleBackColor = true;
            this.ShowWinForm.Click += new System.EventHandler(this.ShowWinForm_Click);
            // 
            // ShowWinMDIForm
            // 
            this.ShowWinMDIForm.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ShowWinMDIForm.Location = new System.Drawing.Point(22, 50);
            this.ShowWinMDIForm.Name = "ShowWinMDIForm";
            this.ShowWinMDIForm.Size = new System.Drawing.Size(286, 29);
            this.ShowWinMDIForm.TabIndex = 1;
            this.ShowWinMDIForm.Text = "ShowWinMDIForm";
            this.ShowWinMDIForm.UseVisualStyleBackColor = true;
            this.ShowWinMDIForm.Click += new System.EventHandler(this.ShowWinMDIForm_Click);
            // 
            // ShowAsCtrl
            // 
            this.ShowAsCtrl.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.ShowAsCtrl.Location = new System.Drawing.Point(22, 85);
            this.ShowAsCtrl.Name = "ShowAsCtrl";
            this.ShowAsCtrl.Size = new System.Drawing.Size(286, 29);
            this.ShowAsCtrl.TabIndex = 2;
            this.ShowAsCtrl.Text = "Show as SubCtrl";
            this.ShowAsCtrl.UseVisualStyleBackColor = true;
            this.ShowAsCtrl.Click += new System.EventHandler(this.ShowAsCtrl_Click);
            // 
            // CLRVsTools
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.ShowAsCtrl);
            this.Controls.Add(this.ShowWinMDIForm);
            this.Controls.Add(this.ShowWinForm);
            this.Name = "CLRVsTools";
            this.Size = new System.Drawing.Size(331, 150);
            this.Load += new System.EventHandler(this.CLRVsTools_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button ShowWinForm;
        private System.Windows.Forms.Button ShowWinMDIForm;
        private System.Windows.Forms.Button ShowAsCtrl;
    }
}
