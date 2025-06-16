using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.UI.Xaml;
using ClipboardToFile.Helpers;
using System;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ClipboardToFile
{
    public partial class App : Application
    {
        private IHost? _host;
        private Window? _window;

        public App()
        {
            InitializeComponent();

            // Handle unhandled exceptions
            UnhandledException += OnUnhandledException;
        }

        /// <summary>
        /// Gets the current app instance in use
        /// </summary>
        public new static App Current => (App)Application.Current;

        /// <summary>
        /// Gets the service provider instance to resolve application services.
        /// </summary>
        public IServiceProvider Services => _host?.Services ?? throw new InvalidOperationException("Services not initialized");

        protected override async void OnLaunched(LaunchActivatedEventArgs args)
        {
            // Build the DI container
            _host = Host.CreateDefaultBuilder()
                .ConfigureServices((context, services) =>
                {
                    // Register all our services
                    services.AddClipboardToFileServices();
                })
                .ConfigureLogging(logging =>
                {
                    logging.AddDebug();
#if DEBUG
                    logging.SetMinimumLevel(LogLevel.Debug);
#endif
                })
                .Build();

            // Start the host
            await _host.StartAsync();

            // Create and activate the main window
            _window = new MainWindow();

            // Handle window closed event for cleanup
            _window.Closed += OnMainWindowClosed;

            _window.Activate();
        }

        private async void OnMainWindowClosed(object sender, WindowEventArgs args)
        {
            // Clean up when main window closes
            await ShutdownAsync();
        }

        private void OnUnhandledException(object sender, Microsoft.UI.Xaml.UnhandledExceptionEventArgs e)
        {
            // Log unhandled exceptions
            var logger = Services.GetService<ILogger<App>>();
            logger?.LogError(e.Exception, "Unhandled exception occurred");

            // Mark as handled to prevent crash (TODO: should we???)
            e.Handled = true;
        }

        private async Task ShutdownAsync()
        {
            try
            {
                if (_host != null)
                {
                    await _host.StopAsync();
                    _host.Dispose();
                    _host = null;
                }
            }
            catch (Exception ex)
            {
                // Log shutdown errors but don't throw
                Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }
    }
}