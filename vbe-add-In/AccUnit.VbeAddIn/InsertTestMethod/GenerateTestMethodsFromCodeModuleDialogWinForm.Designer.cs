namespace AccessCodeLib.AccUnit.VbeAddIn
{
    partial class GenerateTestMethodsFromCodeModuleDialogWinForm
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Methoden");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Eigenschaften");
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GenerateTestMethodsFromCodeModuleDialogWinForm));
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnInsert = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtModuleUnderTest = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtTestclassName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tvwMembers = new AccessCodeLib.AccUnit.VbeAddIn.TriStateCheckBoxesTreeView();
            this.imlTreeView = new System.Windows.Forms.ImageList(this.components);
            this.label5 = new System.Windows.Forms.Label();
            this.txtStateUnderTest = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtExpectedBehaviour = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtMethodNamePreview = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(729, 687);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(151, 48);
            this.btnCancel.TabIndex = 0;
            this.btnCancel.Text = "Abbre&chen";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // btnInsert
            // 
            this.btnInsert.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInsert.Location = new System.Drawing.Point(22, 678);
            this.btnInsert.Name = "btnInsert";
            this.btnInsert.Size = new System.Drawing.Size(138, 45);
            this.btnInsert.TabIndex = 1;
            this.btnInsert.Text = "Ei&nfügen";
            this.btnInsert.UseVisualStyleBackColor = true;
            this.btnInsert.Click += new System.EventHandler(this.btnInsert_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Location = new System.Drawing.Point(14, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(882, 39);
            this.label1.TabIndex = 2;
            this.label1.Text = "Hier können Sie Tests für Members des ausgewählten Moduls einfügen.\r\nWählen Sie d" +
    "azu die Members, für die Testmethoden erstellt werden sollen.";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 73);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(234, 32);
            this.label2.TabIndex = 3;
            this.label2.Text = "Zu testendes Modul:";
            // 
            // txtModuleUnderTest
            // 
            this.txtModuleUnderTest.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtModuleUnderTest.Location = new System.Drawing.Point(285, 69);
            this.txtModuleUnderTest.Name = "txtModuleUnderTest";
            this.txtModuleUnderTest.ReadOnly = true;
            this.txtModuleUnderTest.Size = new System.Drawing.Size(610, 39);
            this.txtModuleUnderTest.TabIndex = 4;
            this.txtModuleUnderTest.Text = "MyModule";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 103);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(237, 32);
            this.label3.TabIndex = 5;
            this.label3.Text = "&Name der Testklasse:";
            // 
            // txtTestclassName
            // 
            this.txtTestclassName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTestclassName.Location = new System.Drawing.Point(285, 99);
            this.txtTestclassName.Name = "txtTestclassName";
            this.txtTestclassName.Size = new System.Drawing.Size(610, 39);
            this.txtTestclassName.TabIndex = 6;
            this.txtTestclassName.Text = "MyModuleTests";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(16, 147);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(567, 32);
            this.label4.TabIndex = 7;
            this.label4.Text = "&Members für die Testklassen erzeugt werden sollen:";
            // 
            // tvwMembers
            // 
            this.tvwMembers.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tvwMembers.AutoCheckChilds = true;
            this.tvwMembers.AutoCheckParents = true;
            this.tvwMembers.ImageIndex = 0;
            this.tvwMembers.ImageList = this.imlTreeView;
            this.tvwMembers.IndeterminateToChecked = true;
            this.tvwMembers.Location = new System.Drawing.Point(20, 182);
            this.tvwMembers.Name = "tvwMembers";
            treeNode1.ImageKey = "Methods";
            treeNode1.Name = "tndMethods";
            treeNode1.SelectedImageKey = "Methods";
            treeNode1.StateImageIndex = 0;
            treeNode1.Text = "Methoden";
            treeNode2.ImageKey = "Properties";
            treeNode2.Name = "tndProperties";
            treeNode2.SelectedImageKey = "Properties";
            treeNode2.StateImageIndex = 0;
            treeNode2.Text = "Eigenschaften";
            this.tvwMembers.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2});
            this.tvwMembers.SelectedImageIndex = 0;
            this.tvwMembers.Size = new System.Drawing.Size(875, 365);
            this.tvwMembers.TabIndex = 8;
            this.tvwMembers.UseTriStateCheckBoxes = true;
            // 
            // imlTreeView
            // 
            this.imlTreeView.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imlTreeView.ImageStream")));
            this.imlTreeView.TransparentColor = System.Drawing.Color.Transparent;
            this.imlTreeView.Images.SetKeyName(0, "Properties");
            this.imlTreeView.Images.SetKeyName(1, "Methods");
            this.imlTreeView.Images.SetKeyName(2, "Member");
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 571);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(222, 32);
            this.label5.TabIndex = 9;
            this.label5.Text = "Getesteter Zu&stand:";
            // 
            // txtStateUnderTest
            // 
            this.txtStateUnderTest.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtStateUnderTest.Location = new System.Drawing.Point(242, 568);
            this.txtStateUnderTest.Name = "txtStateUnderTest";
            this.txtStateUnderTest.Size = new System.Drawing.Size(653, 39);
            this.txtStateUnderTest.TabIndex = 10;
            this.txtStateUnderTest.Text = "StateUnderTest";
            this.txtStateUnderTest.TextChanged += new System.EventHandler(this.txtStateUnderTest_TextChanged);
            // 
            // label6
            // 
            this.label6.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(14, 606);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(236, 32);
            this.label6.TabIndex = 11;
            this.label6.Text = "&Erwartetes Verhalten:";
            // 
            // txtExpectedBehaviour
            // 
            this.txtExpectedBehaviour.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtExpectedBehaviour.Location = new System.Drawing.Point(242, 603);
            this.txtExpectedBehaviour.Name = "txtExpectedBehaviour";
            this.txtExpectedBehaviour.Size = new System.Drawing.Size(653, 39);
            this.txtExpectedBehaviour.TabIndex = 12;
            this.txtExpectedBehaviour.Text = "ExpectedBehaviour";
            this.txtExpectedBehaviour.TextChanged += new System.EventHandler(this.txtExpectedBehaviour_TextChanged);
            // 
            // label7
            // 
            this.label7.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(14, 643);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(299, 32);
            this.label7.TabIndex = 13;
            this.label7.Text = "Name der Testmethode(n):";
            // 
            // txtMethodNamePreview
            // 
            this.txtMethodNamePreview.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtMethodNamePreview.Location = new System.Drawing.Point(242, 639);
            this.txtMethodNamePreview.Name = "txtMethodNamePreview";
            this.txtMethodNamePreview.ReadOnly = true;
            this.txtMethodNamePreview.Size = new System.Drawing.Size(653, 39);
            this.txtMethodNamePreview.TabIndex = 14;
            this.txtMethodNamePreview.Text = "<Member>_StateUnderTest_ExpectedBehaviour";
            // 
            // GenerateTestMethodsFromCodeModuleDialogWinForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(13F, 32F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(910, 735);
            this.Controls.Add(this.txtMethodNamePreview);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.txtExpectedBehaviour);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtStateUnderTest);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.tvwMembers);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.txtTestclassName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtModuleUnderTest);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnInsert);
            this.Controls.Add(this.btnCancel);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MinimumSize = new System.Drawing.Size(505, 540);
            this.Name = "GenerateTestMethodsFromCodeModuleDialogWinForm";
            this.Text = "Testmethoden einfügen";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnInsert;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtModuleUnderTest;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtTestclassName;
        private System.Windows.Forms.Label label4;
        private TriStateCheckBoxesTreeView tvwMembers;
        private System.Windows.Forms.ImageList imlTreeView;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtStateUnderTest;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtExpectedBehaviour;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtMethodNamePreview;
    }
}