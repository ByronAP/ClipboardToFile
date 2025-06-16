using Microsoft.Extensions.DependencyInjection;

namespace ClipboardToFile.Helpers
{
    public static class ServiceCollectionExtensions
    {
        public static IServiceCollection AddClipboardToFileServices(this IServiceCollection services)
        {
            // Core Services

            // Configuration Services

            // ViewModels

            // Add logging
            services.AddLogging();

            return services;
        }
    }
}