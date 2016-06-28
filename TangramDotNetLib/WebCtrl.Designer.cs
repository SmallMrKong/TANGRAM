namespace TangramDotNetLib
{
    partial class WebCtrl
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
            this.webControl = new EO.WebBrowser.WinForm.WebControl();
            this.webView = new EO.WebBrowser.WebView();
            this.SuspendLayout();
            // 
            // webControl
            // 
            this.webControl.BackColor = System.Drawing.Color.White;
            this.webControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webControl.Location = new System.Drawing.Point(0, 0);
            this.webControl.Name = "webControl";
            this.webControl.Size = new System.Drawing.Size(415, 327);
            this.webControl.TabIndex = 0;
            this.webControl.Text = "webCtrl";
            this.webControl.WebView = this.webView;
            // 
            // webView
            // 
            this.webView.Url = "https://www.cloudaddin.com/devx/?redirect=%2Fdevx%2Fchannels%2F8u8muxp6fjdn7xohxkjxa1z5jr__sespjxdogj8ndxtepoxzyot5zh";
            // 
            // WebCtrl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.webControl);
            this.Name = "WebCtrl";
            this.Size = new System.Drawing.Size(415, 327);
            this.ResumeLayout(false);

        }

        #endregion

        private EO.WebBrowser.WinForm.WebControl webControl;
        private EO.WebBrowser.WebView webView;
    }
}
