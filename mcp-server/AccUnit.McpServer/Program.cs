using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using ModelContextProtocol.Protocol;

namespace AccessCodeLib.AccUnit.McpServer
{
    internal static class Program
    {
        private const string ServerName = "AccUnit.McpServer";

        private static async Task Main(string[] args)
        {
            var builder = Host.CreateApplicationBuilder(args);

            // stdout carries the MCP protocol; all log output must go to stderr.
            builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

            builder.Services.AddSingleton<StaComDispatcher>();
            builder.Services.AddSingleton<AccessSession>();
            builder.Services.AddSingleton<TestRunService>();
            builder.Services.AddSingleton<TestClassDraftService>();
            builder.Services.AddSingleton<TestEnvironmentService>();
            builder.Services.AddSingleton<TestClassImportService>();
            builder.Services.AddSingleton<TestModuleExportService>();
            builder.Services.AddSingleton(_ => new SkillInstaller());
            builder.Services
                .AddMcpServer(options =>
                {
                    options.ServerInfo = new Implementation
                    {
                        Name = ServerName,
                        // The server is released together with AccUnit and uses its version.
                        Version = typeof(AccessCodeLib.AccUnit.Interop.AccUnitFactory).Assembly.GetName().Version.ToString(3)
                    };
                    options.ServerInstructions = TestAuthoringGuide.Text;
                })
                .WithStdioServerTransport()
                .WithTools<SessionTools>()
                .WithTools<TestTools>()
                .WithTools<SkillTools>();

            await builder.Build().RunAsync();
        }
    }
}
