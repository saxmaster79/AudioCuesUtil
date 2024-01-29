//setup our DI
using AudioCuesUtil;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.Parsing;
using System.Text.Encodings.Web;
using System.Text.Json;

var cts = new CancellationTokenSource();
Console.CancelKeyPress += (s, e) =>
{
    Console.WriteLine("Canceling...");
    cts.Cancel();
    e.Cancel = true;
};

var builder = new ConfigurationBuilder();
builder
    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
var configuration = builder.Build();


var serviceProvider = new ServiceCollection()
    .AddLogging(loggingBuilder =>
    {
        loggingBuilder.AddConfiguration(configuration.GetSection("Logging"));
        loggingBuilder.AddConsole();
    })
    .AddSingleton<IConfiguration>(configuration)
    .AddCliCommands()
    .AddTransient<IBackupFileReader, BackupFileReader>()
    .AddTransient<IBackupFileWriter, BackupFileWriter>()
    .AddTransient<ExcelWriter>()
    .AddTransient<ExcelReader>()
    .AddSingleton(new JsonSerializerOptions() { PropertyNameCaseInsensitive = true, WriteIndented=true,PropertyNamingPolicy=JsonNamingPolicy.CamelCase, Encoder=JavaScriptEncoder.UnsafeRelaxedJsonEscaping })
    .BuildServiceProvider();

var logger = serviceProvider.GetRequiredService<ILoggerFactory>()
    .CreateLogger<Program>();
logger.LogDebug("Starting application");

var commandLineBuilder = new CommandLineBuilder();



foreach (Command command in serviceProvider.GetServices<Command>())
{
    commandLineBuilder.Command.AddCommand(command);
}

var parser= commandLineBuilder.UseDefaults().Build();
parser.Invoke(args);
logger.LogDebug("All done!");




