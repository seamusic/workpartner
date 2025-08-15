using System;
using Microsoft.Extensions.Configuration;

namespace WorkPartner.Utils
{
	public static class AppSettings
	{
		private static IConfiguration? _configuration;
		public static IConfiguration Configuration => _configuration ??= BuildConfiguration();

		private static IConfiguration BuildConfiguration()
		{
			return new ConfigurationBuilder()
				.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
				.AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
				.AddJsonFile("appsettings.Development.json", optional: true, reloadOnChange: true)
				.AddEnvironmentVariables()
				.Build();
		}
	}
}
