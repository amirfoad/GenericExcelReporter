using GenericExcelReporter.Core.Contracts;
using GenericExcelReporter.Core.Services;
using Microsoft.Extensions.DependencyInjection;

namespace GenericExcelReporter.Core.ServiceConfiguration
{
    public static class ServiceCollectionExtension
    {
        public static IServiceCollection AddExcelService(this IServiceCollection services)
        {
            services.AddScoped(typeof(IExcelService), typeof(ExcelService));
            return services;
        }
    }
}