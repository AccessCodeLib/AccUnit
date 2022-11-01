using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit
{
    internal interface ITransactionManager
    {
        void BeginTrans();
        void Rollback();
    }

    internal class DaoTransactionManager : ITransactionManager
    {
        private readonly DBEngine _dbEngine;

        public DaoTransactionManager(object accessApplication)
        {
            _dbEngine = new AccessApplicationHelper(accessApplication).DBEngine;
        }

        private DBEngine DBEngine
        {
            get { return _dbEngine; }
        }

        public void BeginTrans()
        {
            Logger.Log("Performing auto BeginTrans");
            DBEngine.BeginTrans();
        }

        public void Rollback()
        {
            Logger.Log("Performing auto Rollback");
            DBEngine.Rollback();
        }
    }
}