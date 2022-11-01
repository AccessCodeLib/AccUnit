using System;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Office.Core;

namespace AccessCodeLib.Common.VBIDETools.Commandbar
{
    public static class CommandBarExtensions
    {
        public static CommandBarPopup AddPopup(this CommandBarPopup commandBarPopup)
        {
            using (new BlockLogger())
            {
                return (CommandBarPopup)commandBarPopup.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            }
        }

        public static CommandBarButton AddButton(this CommandBarPopup commandBarPopup)
        {
            using (new BlockLogger())
            {
                return (CommandBarButton)commandBarPopup.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, Type.Missing, true);
            }
        }

        public static CommandBarPopup AddPopup(this CommandBar commandBar)
        {
            using (new BlockLogger())
            {
                return (CommandBarPopup)commandBar.Controls.Add(MsoControlType.msoControlPopup, Type.Missing, Type.Missing, Type.Missing, true);
            }
        }

        public static CommandBarButton AddButton(this CommandBar commandBar, int? before)
        {
            using (new BlockLogger())
            {
                return (CommandBarButton)commandBar.Controls.Add(MsoControlType.msoControlButton, Type.Missing, Type.Missing, before ?? Type.Missing, true);
            }
        }

        public static CommandBarButton AddButton(this CommandBar commandBar)
        {
            using (new BlockLogger())
            {
                return commandBar.AddButton(null);
            }
        }
    }
}