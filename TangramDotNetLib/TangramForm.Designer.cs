namespace TangramDotNetLib
{
    partial class TangramForm
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
            this.TangramPanel = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // TangramPanel
            // 
            this.TangramPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.TangramPanel.Location = new System.Drawing.Point(0, 0);
            this.TangramPanel.Name = "TangramPanel";
            this.TangramPanel.Size = new System.Drawing.Size(713, 474);
            this.TangramPanel.TabIndex = 0;
            // 
            // TangramForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(713, 474);
            this.Controls.Add(this.TangramPanel);
            this.Name = "TangramForm";
            this.Text = "TangramForm";
            this.Load += new System.EventHandler(this.TangramForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel TangramPanel;
    }
}