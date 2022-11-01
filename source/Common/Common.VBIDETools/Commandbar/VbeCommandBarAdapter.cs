using System;
using System.Collections.Generic;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.Common.VBIDETools.Commandbar
{
    public class VbeCommandBarAdapter : IDisposable
    {
        private readonly Stack<CommandBarButtonAndClickHandler> _commandBarButtons = new Stack<CommandBarButtonAndClickHandler>();

        public VbeCommandBarAdapter(VBE vbe)
        {
            VBE = vbe;
        }

        public VBE VBE { get; }

        public CommandBar MenuBar
        {
            get
            {
                const int commandBarIndexMenuBar = 1;
                // name depends on the language (en: "Menu Bar", de: "Menüleiste")
                return GetCommandBar(commandBarIndexMenuBar);
            }
        }

        protected CommandBar CommandBarTools => GetCommandBar("Tools");

        public CommandBar CommandBarView => GetCommandBar("View");

        public CommandBar CommandBarProjectWindow => GetCommandBar("Project Window");

        public CommandBar CommandBarCodeWindow => GetCommandBar("Code Window");

        private CommandBar GetCommandBar(string commandBarName)
        {
            using (new BlockLogger($"by name \"{commandBarName}\""))
            {
                return VBE.CommandBars[commandBarName];
            }
        }

        private CommandBar GetCommandBar(int index)
        {
            using (new BlockLogger($"by index {index}"))
            {
                return VBE.CommandBars[index];
            }
        }

        public static int? GetButtonIndex(CommandBar commandBar, int controlID)
        {
            try
            {
                var foundControl = commandBar.FindControl(MsoControlType.msoControlButton, controlID);
                if (foundControl != null)
                {
                    return foundControl.Index;
                }
            }
            // ReSharper disable EmptyGeneralCatchClause
            catch
            {
                // Don't mind if the control could not be found.
            }
            // ReSharper restore EmptyGeneralCatchClause

            return null;
        }

        public CommandBarButton AddCommandBarButton(CommandBar commandBar, CommandbarButtonData buttonData, _CommandBarButtonEvents_ClickEventHandler handler)
        {
            Logger.Log($"{buttonData.Caption}: {buttonData.Index}");
            var button = AddCommandBarButton(commandBar, buttonData.Index, handler);
            button.Caption = buttonData.Caption;
            button.DescriptionText = buttonData.Description;
            button.FaceId = buttonData.FaceId;
            button.BeginGroup = buttonData.BeginGroup;
            return button;
        }

        public CommandBarButton AddCommandBarButton(CommandBarPopup commandBarPopup, CommandbarButtonData buttonData, _CommandBarButtonEvents_ClickEventHandler handler)
        {
            return AddCommandBarButton(commandBarPopup.CommandBar, buttonData, handler);
        }

        private CommandBarButton AddCommandBarButton(CommandBar commandBar, int? positionBefore, _CommandBarButtonEvents_ClickEventHandler handler)
        {
            var button = commandBar.AddButton(positionBefore);
            button.Click += handler;
            _commandBarButtons.Push(new CommandBarButtonAndClickHandler(button, handler));
            return button;
        }

        public void AddClient(ICommandBarsAdapterClient client)
        {
            using (new BlockLogger(client.GetType().Name))
            {
                client.SubscribeToCommandBarAdapter(this);
            }
        }

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            using (new BlockLogger())
            {
                DisposeUnManagedResources();

                _disposed = true;
            }
        }

        private void DisposeUnManagedResources()
        {
            using (new BlockLogger())
            {

                // issue #77: (http://accunit.access-codelib.net/bugs/view.php?id=77)
                return;
                /*
                try
                {
                    while (_commandBarButtons.Count > 0)
                    {
                        var button = _commandBarButtons.Pop();
                        try
                        {
                            Logger.Log(string.Format("Button: {0}, -ClickEventHandler", button.Button.Caption));
                            button.Button.Click -= button.ClickEventHandler;
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(ex);
                        }
                        try
                        {
                            Logger.Log(string.Format("Button: {0}, Delete", button.Button.Caption));
                            button.Button.Delete();
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(ex);
                        }
                        try
                        {
                            button.Dispose();
                        }
                        catch (Exception ex)
                        {
                            Logger.Log(ex);
                        }
                    }
                    Logger.Log("_commandBarButtons.Clear");
                    _commandBarButtons.Clear();
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }
                */
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~VbeCommandBarAdapter()
        {
            Dispose(false);
        }

        #endregion

        private sealed class CommandBarButtonAndClickHandler : IDisposable
        {
            private CommandBarButton _commandBarButton;
            private readonly _CommandBarButtonEvents_ClickEventHandler _clickEventHandler;

            public CommandBarButtonAndClickHandler(CommandBarButton commandBarButton, _CommandBarButtonEvents_ClickEventHandler clickEventHandler)
            {
                _commandBarButton = commandBarButton;
                _clickEventHandler = clickEventHandler;
            }

            public CommandBarButton Button
            {
                get { return _commandBarButton; }
            }

            public _CommandBarButtonEvents_ClickEventHandler ClickEventHandler
            {
                get { return _clickEventHandler; }
            }
                    
            #region IDisposable Support

            bool _disposed;

            private void Dispose(bool disposing)
            {
                if (_disposed)
                    return;

                using (new BlockLogger())
                {
                    if (disposing)
                    {
                    }

                    DisposeUnManagedResources();

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    _disposed = true;
                }
            }

            private void DisposeUnManagedResources()
            {
                using (new BlockLogger())
                {
                    _commandBarButton = null;
                }
            }

            public void Dispose()
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }

            ~CommandBarButtonAndClickHandler()
            {
                Dispose(false);
            }

        #endregion
        }
    }
}
