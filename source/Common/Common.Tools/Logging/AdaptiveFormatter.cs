namespace AccessCodeLib.Common.Tools.Logging
{
    internal class AdaptiveFormatter
    {
        private readonly string _metaFormat;

        public AdaptiveFormatter(string metaFormat)
        {
            _metaFormat = metaFormat;
        }

        public string MetaFormat
        {
            get { return _metaFormat; }
        }

        public string GetFormattedInfo(object info)
        {
            var actualInfoLength = info.ToString().Length;
            if (actualInfoLength > CurrentInfoLength)
            {
                CurrentInfoLength = actualInfoLength;
            }
            var format = string.Format(MetaFormat, CurrentInfoLength);
            //Logger.LogRaw(string.Format("Actual format: \"{0}\", actual info: \"{1}\"", format, info));
            return string.Format(format, info);
        }

        public int CurrentInfoLength { get; set; }
    }
}