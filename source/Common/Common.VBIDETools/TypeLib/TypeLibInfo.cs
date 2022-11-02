using TLI = TypeLibInformation;

namespace AccessCodeLib.Common.VBIDETools.TypeLib
{
    public class TypeLibInfo
    {
        private readonly TLI.TypeLibInfo _typeLibInfo;
        private Constants _constants;

        public TypeLibInfo(TLI.TypeLibInfo typeLibInfo)
        {
            _typeLibInfo = typeLibInfo;
        }

        public TypeLibInfo(string libFileName)
        {
            var tliApp = new TLI.TLIApplication();
            _typeLibInfo = tliApp.TypeLibInfoFromFile(libFileName);
        }

        public string Name
        {
            get { return _typeLibInfo.Name; }
        }

        public Constants Constants
        {
            get
            {
                if (_constants == null)
                {
                    _constants = new Constants();
                    ReadConstants();
                }
                return _constants;
            }
        }

        private void ReadConstants()
        {
            foreach (TLI.ConstantInfo constantInfo in _typeLibInfo.Constants)
            {
                _constants.Add(constantInfo);
            }
        }

    }
}
