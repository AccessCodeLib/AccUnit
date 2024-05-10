using System;
using System.Windows.Forms;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class MessageEventArgs : EventArgs
    {
        private readonly string _message;
        private readonly MessageBoxButtons _buttons;
        private readonly MessageBoxIcon _icon;
        private readonly MessageBoxDefaultButton _defaultButton;

        public MessageEventArgs(string message,
                                MessageBoxButtons buttons,
                                MessageBoxIcon icon,
                                MessageBoxDefaultButton defaultButton)
        {
            _message = message;
            _buttons = buttons;
            _icon = icon;
            _defaultButton = defaultButton;
        }

        public string Message
        {
            get { return _message; }
        }

        public MessageBoxButtons Buttons
        {
            get { return _buttons; }
        }

        public MessageBoxIcon Icon
        {
            get { return _icon; }
        }

        public MessageBoxDefaultButton DefaultButton
        {
            get { return _defaultButton; }
        }

        public bool MessageDisplayed { get; set; }
    }
}
