using System.Runtime.InteropServices;

namespace AccessCodeLib.OpenAI.VbeAddIn
{
    [ComVisible(true)]
    [Guid("C0FD45FA-2A0C-4D52-B3ED-731754A050EB")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccessCodeLib.OpenAI.AddInManager")]
    public class AddInManagerBridge : IAddInManagerBridge
    {
        public delegate void TestCodeBuilderFactoryEventHandler(out ITestCodeBuilderFactory testCodeBuilderFactory);
        public event TestCodeBuilderFactoryEventHandler TestCodeBuilderFactoryRequest;

        
        public delegate void HostApplicationInitializedEventHandler(object hostapplication);

        public event HostApplicationInitializedEventHandler HostApplicationInitialized;

        public object Application
        {
            set
            {
                HostApplicationInitialized?.Invoke(value);
            }
        }

        public ITestCodeBuilderFactory TestCodeBuilderFactory()
        {
            ITestCodeBuilderFactory builderFactory = null;
            TestCodeBuilderFactoryRequest?.Invoke(out builderFactory);
            return builderFactory;
        }

    }

    [ComVisible(true)]
    [Guid("0EEEA3E7-68D6-49BA-8536-572E69CCCEF0")]
    public interface IAddInManagerBridge
    {
        object Application { set; }
        ITestCodeBuilderFactory TestCodeBuilderFactory();
    }

}
