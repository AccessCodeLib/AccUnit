using AccessCodeLib.Common.Tools.Files;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;

namespace AccessCodeLib.Common.VBIDETools
{
    public interface IAccessApplicationHelper
    {
        CurrentDb CurrentDb { get; }
        DBEngine DBEngine { get; }
        DoCmd DoCmd { get; }
        bool IsCompiled { get; }

        object GetOption(string optionName);
        void LoadFromText(AccessApplicationHelper.AcObjectType objectType, string objectName, string fileName);
        void RunCommand(AccessApplicationHelper.AcCommand command);
        void SaveAsText(AccessApplicationHelper.AcObjectType objectType, string objectName, string fileName);
        void SetOption(string optionName, object optionValue);
    }

    public class AccessApplicationHelper : OfficeApplicationHelper, IAccessApplicationHelper
    {
        public enum AcCommand { AcCmdCompileAndSaveAllModules = 126 };

        public enum AcObjectType { AcDefault = -1, AcTable = 0, AcQuery = 1, AcForm = 2, AcReport = 3, AcMacro = 4, AcModule = 5 };

        public AccessApplicationHelper(object application)
            : base(application)
        {
        }

        public DBEngine DBEngine { get { return new DBEngine(Application); } }
        public CurrentDb CurrentDb { get { return new CurrentDb(Application); } }
        public bool IsCompiled { get { return (bool)InvocationHelper.InvokePropertyGet(AccessInvokeStrings.Application.IsCompiled); } }

        public object GetOption(string optionName)
        {
            return InvocationHelper.InvokeMethod(AccessInvokeStrings.Application.GetOption, new object[] { optionName });
        }

        public void SetOption(string optionName, object optionValue)
        {
            InvocationHelper.InvokeMethod(AccessInvokeStrings.Application.SetOption, new[] { optionName, optionValue });
        }

        public void RunCommand(AcCommand command)
        {
            InvocationHelper.InvokeMethod(AccessInvokeStrings.Application.RunCommand, new object[] { (Int32)command });
        }

        protected override VBProject GetCheckedVbProject()
        {
            using (new BlockLogger())
            {
                var activeVBProject = VBE.ActiveVBProject;

                string currentDbName;
                try
                {
                    currentDbName = FileTools.ConvertPathToUNC(CurrentDb.Name);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                    return activeVBProject;
                }

                if (!VbProjectMatchWithFullName(activeVBProject, currentDbName))
                {
                    activeVBProject = FindVbProjectEqualsFullName(currentDbName);
                }
                return activeVBProject;

            }
        }

        public void SaveAsText(AcObjectType objectType, string objectName, string fileName)
        {
            Logger.Log(string.Format("objectType: {0}, objectName: {1}, fileName: {2}", (Int32)objectType, objectName, fileName));
            InvocationHelper.InvokeMethod(AccessInvokeStrings.Application.SaveAsText, new object[] { (Int32)objectType, objectName, fileName });
        }

        public void LoadFromText(AcObjectType objectType, string objectName, string fileName)
        {
            Logger.Log(string.Format("objectType: {0}, objectName: {1}, fileName: {2}", (Int32)objectType, objectName, fileName));
            InvocationHelper.InvokeMethod(AccessInvokeStrings.Application.LoadFromText, new object[] { (Int32)objectType, objectName, fileName });
        }

        public DoCmd DoCmd { get { return new DoCmd(Application); } }

    }

    public class CurrentDb
    {
        private readonly object _currentDb;
        private readonly InvocationHelper _invocationHelper;

        public CurrentDb(object accessApplication)
        {
            _currentDb = new InvocationHelper(accessApplication).InvokeMethod(AccessInvokeStrings.Application.CurrentDb, null);
            if (_currentDb == null)
                throw new InvalidOperationException("No database loaded into this Access.Application.");
            _invocationHelper = new InvocationHelper(_currentDb);
        }

        public object Instance { get { return _currentDb; } }

        public string Name { get { return (string)_invocationHelper.InvokePropertyGet(AccessInvokeStrings.DAO.Database.Name); } }
    }

    public class DBEngine
    {
        private readonly object _dbEngine;
        private readonly InvocationHelper _invocationHelper;

        public DBEngine(object accessApplication)
        {
            _dbEngine = new InvocationHelper(accessApplication).InvokePropertyGet(AccessInvokeStrings.Application.DbEngine);
            _invocationHelper = new InvocationHelper(_dbEngine);
        }

        public object Instance { get { return _dbEngine; } }

        public void BeginTrans()
        {
            _invocationHelper.InvokeMethod(AccessInvokeStrings.DBEngine.BeginTrans, null);
        }

        public void Rollback()
        {
            _invocationHelper.InvokeMethod(AccessInvokeStrings.DBEngine.Rollback, null);
        }
    }

    public class DoCmd
    {
        private readonly object _doCmd;
        private readonly InvocationHelper _invocationHelper;

        public DoCmd(object accessApplication)
        {
            _doCmd = new InvocationHelper(accessApplication).InvokePropertyGet(AccessInvokeStrings.Application.DoCmd);
            _invocationHelper = new InvocationHelper(_doCmd);
        }

        public object Instance { get { return _doCmd; } }

        public void DeleteObject(AccessApplicationHelper.AcObjectType objectType, string objectName)
        {
            _invocationHelper.InvokeMethod(AccessInvokeStrings.DoCmdStrings.DeleteObject, new object[] { (Int32)objectType, objectName });
        }

    }

}
