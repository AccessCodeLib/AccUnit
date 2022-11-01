using System;

namespace AccessCodeLib.Common.VBIDETools
{
    
    public interface IOfficeInvokeStrings
    {
        IOfficeApplicationInvokeStrings Application { get; }
    }

    public interface IOfficeApplicationInvokeStrings
    {
        string Name { get; }
        string Run { get; }
        string ActiveDocument { get; }
        IOfficeActiveDocumentInvokeStrings ActiveDocumentInvokeStrings { get; }
    }

    public interface IOfficeActiveDocumentInvokeStrings
    {
        string FullName { get; }
    }

    public class AccessInvokeStrings : IOfficeInvokeStrings
    {
        IOfficeApplicationInvokeStrings IOfficeInvokeStrings.Application { get { return ApplicationInvokeStrings; } }
        
        public AccessInvokeStrings()
        {
            ApplicationInvokeStrings = new Application();
        }

        public Application ApplicationInvokeStrings { get; private set; }

        public class Application : IOfficeApplicationInvokeStrings
        {
            string IOfficeApplicationInvokeStrings.Name { get { return Name; } }
            string IOfficeApplicationInvokeStrings.Run { get { return Run; } }
            string IOfficeApplicationInvokeStrings.ActiveDocument { get { return CurrentDb; } }
            IOfficeActiveDocumentInvokeStrings IOfficeApplicationInvokeStrings.ActiveDocumentInvokeStrings {get { return new DAO.Database(); }}

            public const string InvokeString = "Application";
            public const string Name = "Name";
            public const string Run = "Run";
            public const string CurrentDb = "CurrentDb";
            public const string GetOption = "GetOption";
            public const string SetOption = "SetOption";
        
            public const string RunCommand = "RunCommand";
            public const string IsCompiled = "IsCompiled";
            public const string DbEngine = "DBEngine";

            public const string SaveAsText = "SaveAsText";
            public const string LoadFromText = "LoadFromText";
            public const string DoCmd = "DoCmd";

        }

        public class DAO
        {
            public class Database : IOfficeActiveDocumentInvokeStrings
            {
                string IOfficeActiveDocumentInvokeStrings.FullName { get { return Name; } }
                
                public const string Name = "Name";
            }
        }

        public class DBEngine
        {
            public const string BeginTrans = "BeginTrans";
            public const string Rollback = "Rollback";
        }

        public class DoCmdStrings
        {
            public const string DeleteObject = "DeleteObject";
        }

    }

    public class ExcelInvokeStrings : IOfficeInvokeStrings
    {
        IOfficeApplicationInvokeStrings IOfficeInvokeStrings.Application { get { return new Application(); } }

        public class Application : IOfficeApplicationInvokeStrings
        {
            string IOfficeApplicationInvokeStrings.Name { get { return Name; } }
            string IOfficeApplicationInvokeStrings.Run { get { return Run; } }
            string IOfficeApplicationInvokeStrings.ActiveDocument { get { return ActiveWorkBook; } }
            IOfficeActiveDocumentInvokeStrings IOfficeApplicationInvokeStrings.ActiveDocumentInvokeStrings { get { return new WorkBook(); } }

            public const string InvokeString = "Application";
            public const string Name = "Name";
            public const string Run = "Run";
            public const string ActiveWorkBook = "ActiveWorkBook";
        }

        public class WorkBook : IOfficeActiveDocumentInvokeStrings
        {
            string IOfficeActiveDocumentInvokeStrings.FullName { get { return FullName; } }
            public const string FullName = "FullName";
        }

    }

    public class WordInvokeStrings : IOfficeInvokeStrings
    {
        IOfficeApplicationInvokeStrings IOfficeInvokeStrings.Application { get { return new Application(); } }

        public class Application : IOfficeApplicationInvokeStrings
        {
            string IOfficeApplicationInvokeStrings.Name { get { return Name; } }
            string IOfficeApplicationInvokeStrings.Run { get { return Run; } }
            string IOfficeApplicationInvokeStrings.ActiveDocument { get { return ActiveDocument; } }
            IOfficeActiveDocumentInvokeStrings IOfficeApplicationInvokeStrings.ActiveDocumentInvokeStrings { get { return new WorkBook(); } }

            public const string InvokeString = "Application";
            public const string Name = "Name";
            public const string Run = "Run";
            public const string ActiveDocument = "ActiveDocument";
        }

        public class WorkBook : IOfficeActiveDocumentInvokeStrings
        {
            string IOfficeActiveDocumentInvokeStrings.FullName { get { return FullName; } }
            public const string FullName = "FullName";
        }
    }

    public class PowerPointInvokeStrings : IOfficeInvokeStrings
    {
        IOfficeApplicationInvokeStrings IOfficeInvokeStrings.Application { get { return new Application(); } }

        public class Application : IOfficeApplicationInvokeStrings
        {
            string IOfficeApplicationInvokeStrings.Name { get { return Name; } }
            string IOfficeApplicationInvokeStrings.Run { get { return Run; } }
            string IOfficeApplicationInvokeStrings.ActiveDocument { get { return ActivePresentation; } }
            IOfficeActiveDocumentInvokeStrings IOfficeApplicationInvokeStrings.ActiveDocumentInvokeStrings { get { return new WorkBook(); } }

            public const string InvokeString = "Application";
            public const string Name = "Name";
            public const string Run = "Run";
            public const string ActivePresentation = "ActivePresentation";
        }

        public class WorkBook : IOfficeActiveDocumentInvokeStrings
        {
            string IOfficeActiveDocumentInvokeStrings.FullName { get { return FullName; } }
            public const string FullName = "FullName";
        }
    }

    public static class OfficeInvokeNamesBuilder
    {
        public static IOfficeApplicationInvokeStrings OfficeApplicationInvokeStrings(string applicationName)
        {
            return GetOfficeInvokeNames(applicationName).Application;
        }

        public static IOfficeInvokeStrings GetOfficeInvokeNames(string applicationName)
        {
            switch (applicationName)
            {
                case "Microsoft Access":
                    return new AccessInvokeStrings();
                case "Microsoft Excel":
                    return new ExcelInvokeStrings();
                case "Microsoft Word":
                    return new AccessInvokeStrings();
                case "Microsoft PowerPoint":
                    return new AccessInvokeStrings();
            }
            throw new NotSupportedException(applicationName);
        }
    }

}
