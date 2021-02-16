using Microsoft.Extensions.DependencyInjection;
using System;

namespace SimplyExcel.NET
{
    public static class Installer
    {
        public static IServiceCollection AddSimplyExcel(this IServiceCollection services)
        {
            services.AddTransient<ISimplyExcel, SimplyExcel>();
            return services;
        }

        public static IServiceCollection AddSimplyExcel<TImplementation>(this IServiceCollection services) where TImplementation : class
        {
            services.AddTransient(typeof(ISimplyExcel), typeof(TImplementation));
            return services;
        }
    }
}
