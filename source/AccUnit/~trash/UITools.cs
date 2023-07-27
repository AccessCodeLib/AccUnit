/*
using AccessCodeLib.Common.Tools.Logging;
using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.Common
{
    public static class UITools
    {
        public static void ShowException(Exception exception)
        {
            Logger.Log(exception, 1);
            var message = new StringBuilder("");

            AppendExceptionInfo(message, exception);

            MessageBox.Show(message.ToString(), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);

            var modifierKeys = Control.ModifierKeys;
            if ((modifierKeys & Keys.ShiftKey) <= 0) return;
            Clipboard.SetText(message.ToString());
            MessageBox.Show(MessageStrings.UITools_ShowException_message_has_been_copied_to_the_clipboard,
                            Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void AppendExceptionInfo(StringBuilder message, Exception exception)
        {
            message.AppendLine(exception.Message);

#if DEBUG
            message.AppendFormat("\nStackTrace:\n{0}\n", exception.StackTrace);
#endif
            if (exception.InnerException == null) return;
            message.AppendLine("\nInnerException:");
            AppendExceptionInfo(message, exception.InnerException);
        }


        public static DialogResult ShowMessage(string message, MessageBoxButtons buttons = MessageBoxButtons.OK,
                                               MessageBoxIcon icon = MessageBoxIcon.Information,
                                               MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
        {
            return MessageBox.Show(message, Application.ProductName, buttons, icon, defaultButton);
        }

        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            // see: http://www.csharp-examples.net/inputbox/

            var form = new Form();
            var label = new Label();
            var textBox = new TextBox();
            var buttonOk = new Button();
            var buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

            buttonOk.Text = UserControls.InputBoxButtonOkCaption;
            buttonCancel.Text = UserControls.InputBoxButtonCancelCaption;
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;

            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);

            label.AutoSize = true;
            textBox.Anchor |= AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;

            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;

            var dialogResult = form.ShowDialog();
            value = textBox.Text;
            return dialogResult;
        }

        public static Icon ConvertBitmapToIcon(Bitmap bmp)
        {
            return Icon.FromHandle(bmp.GetHicon());
        }
    }
}
*/