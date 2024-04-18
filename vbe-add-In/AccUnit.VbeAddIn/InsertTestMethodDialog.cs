using System;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public partial class InsertTestMethodDialog : Form
    {
        public event EventHandler<CommitInsertTestMethodEventArgs> CommitMethodName;

        public InsertTestMethodDialog()
        {
            InitializeComponent();
        }

        private void OkButtonClick(object sender, EventArgs e)
        {
            if (CommitMethodName != null)
            {
                var methodUnderTest = methodUnderTestTextBox.Text;
                var stateUnderTest = stateUnderTestTextBox.Text;
                var expectedBehaviour = expectedBehaviourTextBox.Text;
                CommitMethodName(this, new CommitInsertTestMethodEventArgs(methodUnderTest, stateUnderTest, expectedBehaviour));
            }
            Close();
        }
    }
}
