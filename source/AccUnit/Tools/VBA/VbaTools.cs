namespace AccessCodeLib.AccUnit.Tools.VBA
{
    public static class VbaTools
    {
        private static VbaConstantsDictionary _VbaConstantsDictionary;

        public static VbaConstantsDictionary ConstantsDictionary
        {
            get
            {
                if (_VbaConstantsDictionary is null)
                {
                    _VbaConstantsDictionary = new VbaConstantsDictionary();
                }
                return _VbaConstantsDictionary;
            }
        }
    }

}
