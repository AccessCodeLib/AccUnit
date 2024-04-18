namespace AccessCodeLib.AccUnit.VbeAddIn
{
    partial class InsertTestMethodDialog
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
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.methodUnderTestTextBox = new System.Windows.Forms.TextBox();
            this.expectedBehaviourTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.stateUnderTestTextBox = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // cancelButton
            // 
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(275, 102);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(87, 27);
            this.cancelButton.TabIndex = 5;
            this.cancelButton.Text = "&Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            // 
            // okButton
            // 
            this.okButton.Location = new System.Drawing.Point(12, 102);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(87, 27);
            this.okButton.TabIndex = 4;
            this.okButton.Text = "&OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.OkButtonClick);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 15);
            this.label1.TabIndex = 2;
            this.label1.Text = "&Method under test";
            // 
            // methodUnderTestTextBox
            // 
            this.methodUnderTestTextBox.Location = new System.Drawing.Point(132, 7);
            this.methodUnderTestTextBox.Name = "methodUnderTestTextBox";
            this.methodUnderTestTextBox.Size = new System.Drawing.Size(230, 23);
            this.methodUnderTestTextBox.TabIndex = 1;
            // 
            // expectedBehaviourTextBox
            // 
            this.expectedBehaviourTextBox.Location = new System.Drawing.Point(132, 66);
            this.expectedBehaviourTextBox.Name = "expectedBehaviourTextBox";
            this.expectedBehaviourTextBox.Size = new System.Drawing.Size(230, 23);
            this.expectedBehaviourTextBox.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 69);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(110, 15);
            this.label2.TabIndex = 4;
            this.label2.Text = "&Expected behaviour";
            // 
            // stateUnderTestTextBox
            // 
            this.stateUnderTestTextBox.Location = new System.Drawing.Point(132, 36);
            this.stateUnderTestTextBox.Name = "stateUnderTestTextBox";
            this.stateUnderTestTextBox.Size = new System.Drawing.Size(230, 23);
            this.stateUnderTestTextBox.TabIndex = 2;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(8, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(89, 15);
            this.label3.TabIndex = 6;
            this.label3.Text = "&State under test";
            // 
            // InsertTestMethodDialog
            // 
            this.AcceptButton = this.okButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(372, 137);
            this.Controls.Add(this.stateUnderTestTextBox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.expectedBehaviourTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.methodUnderTestTextBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.okButton);
            this.Controls.Add(this.cancelButton);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "InsertTestMethodDialog";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Insert new test method";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox methodUnderTestTextBox;
        private System.Windows.Forms.TextBox expectedBehaviourTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox stateUnderTestTextBox;
        private System.Windows.Forms.Label label3;
    }
}
