namespace AccessCodeLib.AccUnit.VbeAddIn
{
    partial class SelectListControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SelectListControl));
            this.selectAllCheckBox = new System.Windows.Forms.CheckBox();
            this.RefreshButton = new System.Windows.Forms.Button();
            this.CommitButton = new System.Windows.Forms.Button();
            this.itemListBox = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // selectAllCheckBox
            // 
            resources.ApplyResources(this.selectAllCheckBox, "selectAllCheckBox");
            this.selectAllCheckBox.Name = "selectAllCheckBox";
            this.selectAllCheckBox.ThreeState = true;
            this.selectAllCheckBox.UseVisualStyleBackColor = true;
            this.selectAllCheckBox.CheckStateChanged += new System.EventHandler(this.SelectAllCheckBoxCheckStateChanged);
            // 
            // RefreshButton
            // 
            resources.ApplyResources(this.RefreshButton, "RefreshButton");
            this.RefreshButton.CausesValidation = false;
            this.RefreshButton.Image = global::AccessCodeLib.AccUnit.VbeAddIn.Properties.Resources.refresh_green_16x16;
            this.RefreshButton.Name = "RefreshButton";
            this.RefreshButton.UseVisualStyleBackColor = true;
            this.RefreshButton.Click += new System.EventHandler(this.RefreshButtonClick);
            // 
            // CommitButton
            // 
            resources.ApplyResources(this.CommitButton, "CommitButton");
            this.CommitButton.Name = "CommitButton";
            this.CommitButton.UseVisualStyleBackColor = true;
            this.CommitButton.Click += new System.EventHandler(this.CommitSelectedTestsButtonClick);
            // 
            // itemListBox
            // 
            resources.ApplyResources(this.itemListBox, "itemListBox");
            this.itemListBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.itemListBox.CheckOnClick = true;
            this.itemListBox.FormattingEnabled = true;
            this.itemListBox.MultiColumn = true;
            this.itemListBox.Name = "itemListBox";
            this.itemListBox.Sorted = true;
            this.itemListBox.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.ItemListBoxItemCheck);
            // 
            // SelectListControl
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.itemListBox);
            this.Controls.Add(this.selectAllCheckBox);
            this.Controls.Add(this.RefreshButton);
            this.Controls.Add(this.CommitButton);
            this.Name = "SelectListControl";
            this.Resize += new System.EventHandler(this.SelectListControlResize);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        protected internal System.Windows.Forms.Button CommitButton;
        protected internal System.Windows.Forms.Button RefreshButton;
        protected internal System.Windows.Forms.CheckBox selectAllCheckBox;
        protected internal System.Windows.Forms.CheckedListBox itemListBox;

    }
}
