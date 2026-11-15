using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace AccessCodeLib.AccUnit.McpServer
{
    internal static class Program
    {
        private static async Task Main(string[] args)
        {
            var builder = Host.CreateApplicationBuilder(args);

            // stdout carries the MCP protocol; all log output must go to stderr.
            builder.Logging.AddConsole(options => options.LogToStandardErrorThreshold = LogLevel.Trace);

            builder.Services.AddSingleton<StaComDispatcher>();
            builder.Services.AddSingleton<AccessSession>();
            builder.Services
                .AddMcpServer()
                .WithStdioServerTransport()
                .WithTools<SessionTools>();

            await builder.Build().RunAsync();
        }
    }
}
