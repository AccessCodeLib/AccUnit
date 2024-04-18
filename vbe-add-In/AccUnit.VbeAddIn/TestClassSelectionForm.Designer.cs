namespace AccessCodeLib.AccUnit.VbeAddIn
{
    sealed partial class TestClassSelectionForm
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
            this.testListUserControl = new TestListControl();
            this.SuspendLayout();
            // 
            // testListUserControl
            // 
            this.testListUserControl.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.testListUserControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.testListUserControl.Location = new System.Drawing.Point(0, 0);
            this.testListUserControl.Name = "testListUserControl";
            this.testListUserControl.Size = new System.Drawing.Size(334, 309);
            this.testListUserControl.TabIndex = 0;
            // 
            // TestClassSelectionForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(334, 309);
            this.Controls.Add(this.testListUserControl);
            this.Name = "TestClassSelectionForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Select Tests";
            this.ResumeLayout(false);

        }

        #endregion

        private TestListControl testListUserControl;
    }
}