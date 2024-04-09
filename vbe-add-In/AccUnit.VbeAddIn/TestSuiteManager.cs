using System;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    class TestSuiteManager : IDisposable
    {
        public delegate void TestSuiteInitializedEventHandler(ITestSuite suite);
        public event TestSuiteInitializedEventHandler TestSuiteInitialized;

        /*
        public delegate void TestCountChangedEventHandler(int number);
        public event TestCountChangedEventHandler TestCountChanged;
        */

        private IVBATestSuite _vbaTestSuite;
        private readonly ITestResultReporter _testResultReporter = new TestResultReporter();
        public ITestResultReporter TestResultReporter
        {
            get { return _testResultReporter; }
        }

        private OfficeApplicationHelper _officeApplicationHelper;
        public OfficeApplicationHelper OfficeApplicationHelper
        {
            get { return _officeApplicationHelper; }
            set 
            { 
                _officeApplicationHelper = value;
            }
        }

        public IVBATestSuite TestSuite
        {
            get
            {
                if (_vbaTestSuite == null)
                {
                    InitTestSuite();
                }
                return _vbaTestSuite;
            }
        }

        public VBProject ActiveVBProject
        {
            get
            {
                return ((VBATestSuite)TestSuite).ActiveVBProject;
            }
        }

        private void InitTestSuite()
        {
            using (new BlockLogger())
            {
                try
                {
                    _vbaTestSuite = CreateVbaTestSuite(OfficeApplicationHelper);
                    _vbaTestSuite?.AppendTestResultReporter(_testResultReporter);
                }
                catch (Exception ex)
                {
                    UITools.ShowException(ex);
                }
                finally
                {
                    TestSuiteInitialized?.Invoke(_vbaTestSuite);
                }
            }
        }

        private LoggerForm _loggerForm; 

        private IVBATestSuite CreateVbaTestSuite(OfficeApplicationHelper applicationHelper)
        {
            using (new BlockLogger())
            {
                IVBATestSuite vbaTestSuite;
                var accUnitFactory = new Interop.AccUnitFactory();
                if (applicationHelper is AccessApplicationHelper)
                {
                    Logger.Log("Access application");
                    vbaTestSuite = accUnitFactory.AccessTestSuite(applicationHelper);
                }
                else
                {
                    vbaTestSuite = accUnitFactory.VBATestSuite(applicationHelper);
                }

                if (_loggerForm == null)
                    _loggerForm = new LoggerForm();

                _loggerForm.Visible = true;  
                vbaTestSuite.AppendTestResultReporter(new LoggerFormReporter(_loggerForm));
                return vbaTestSuite;
            }
        }

        #region IDisposable Support

        bool _disposed;

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                DisposeManagedResources();
            }

            DisposeUnmanagedResources();
            _disposed = true;
        }

        private void DisposeUnmanagedResources()
        {
            OfficeApplicationHelper = null;
        }

        private void DisposeManagedResources()
        {
            DisposeVbaTestSuite();
        }

        private void DisposeVbaTestSuite()
        {
            if (_vbaTestSuite == null)
                return;

            using (new BlockLogger())
            {
                try
                {
                    _vbaTestSuite.Dispose();
                    Logger.Log("_vbaTestSuite disposed");
                }
                catch (Exception exception)
                {
                    Logger.Log(exception);
                }
                finally
                {
                    _vbaTestSuite = null;
                }
            }   
        }

        public void Dispose()
        {
            using (new BlockLogger())
            {
                Dispose(true);
                GC.SuppressFinalize(this);
            }
        }

        ~TestSuiteManager()
        {
            Dispose(false);
        }
        #endregion

    }
}
