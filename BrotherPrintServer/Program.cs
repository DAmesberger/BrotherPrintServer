using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Specialized;
using EmbedIO.Actions;
using EmbedIO.Files;
using EmbedIO.Security;
using EmbedIO.WebApi;
using Swan;
using Swan.Logging;
using System.Threading;
using System.Diagnostics;
using EmbedIO;
using EmbedIO.Routing;
using EmbedIO.Utilities;

namespace BrotherPrintServer
{
    enum Cmd
    {
        GetPrinters,
        GetTemplates,
        Print
    }

    class Program
    {
        static void Main(string[] args)
        {
            var url = args.Length > 0 ? args[0] : "http://*:60024";

            Console.Write("Updating printer settings...");
            Task.WaitAll(Brother.UpdateAsync());
            Console.WriteLine("done");

            using (var cts = new CancellationTokenSource())
            {
                Task.WaitAll(
                    RunWebServerAsync(url, cts.Token),
                    Task.CompletedTask,
                    WaitForUserBreakAsync(cts.Cancel));
            }

            // Clean up
            "Bye".Info(nameof(Program));
            Terminal.Flush();

            Console.WriteLine("Press any key to exit.");
            WaitForKeypress();
        }

        // Gets the local path of shared files.
        // When debugging, take them directly from source so we can edit and reload.
        // Otherwise, take them from the deployment directory.
        public static string HtmlRootPath
        {
            get
            {
                var assemblyPath = Path.GetDirectoryName(typeof(Program).Assembly.Location);

#if DEBUG
                return Path.Combine(Directory.GetParent(assemblyPath).Parent.Parent.FullName, "html");
#else
                return Path.Combine(assemblyPath, "html");
#endif
            }
        }

        // Create and configure our web server.
        private static WebServer CreateWebServer(string url)
        {
#pragma warning disable CA2000 // Call Dispose on object - this is a factory method.
            var server = new WebServer(o => o
                    .WithUrlPrefix(url)
                    .WithMode(HttpListenerMode.EmbedIO))
                .WithIPBanning(o => o
                    .WithMaxRequestsPerSecond()
                    .WithRegexRules("HTTP exception 404"))
                .WithLocalSessionManager()
                .WithCors()
                .WithWebApi("/api", m => m
                    .WithController<BrotherController>());
            // Listen for state changes.
            server.StateChanged += (s, e) => $"WebServer New State - {e.NewState}".Info();

            return server;
#pragma warning restore CA2000
        }

        // Create and run a web server.
        private static async Task RunWebServerAsync(string url, CancellationToken cancellationToken)
        {
            using var server = CreateWebServer(url);
            await server.RunAsync(cancellationToken).ConfigureAwait(false);
        }

        // Open the default browser on the web server's home page.
#pragma warning disable CA1801 // Unused parameter
        private static async Task ShowBrowserAsync(string url, CancellationToken cancellationToken)
#pragma warning restore CA1801
        {
            // Be sure to run in parallel.
            await Task.Yield();

            // Fire up the browser to show the content!
            using var browser = new Process
            {
                StartInfo = new ProcessStartInfo(url)
                {
                    UseShellExecute = true,
                },
            };
            browser.Start();
        }

        // Prompt the user to press any key; when a key is next pressed,
        // call the specified action to cancel operations.
        private static async Task WaitForUserBreakAsync(Action cancel)
        {
            // Be sure to run in parallel.
            await Task.Yield();

            "Press any key to stop the web server.".Info(nameof(Program));
            WaitForKeypress();
            "Stopping...".Info(nameof(Program));
            cancel();
        }

        // Clear the console input buffer and wait for a keypress
        private static void WaitForKeypress()
        {
            while (Console.KeyAvailable)
                Console.ReadKey(true);

            Console.ReadKey(true);
        }
    }
}
