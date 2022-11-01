namespace AccessCodeLib.Common.VBIDETools.Commandbar
{
    public class CommandbarButtonData
    {
        public CommandbarButtonData()
        {
            FaceId = 0;
            BeginGroup = false;
        }

        public CommandbarButtonData(string caption)
        {
            Caption = caption;
        }

        public string Caption { get; set; }
        public string Description { get; set; }
        public int FaceId { get; set; }
        public bool BeginGroup { get; set; }
        public int? Index { get; set; }
        public string Tag { get; set; }
    }
}