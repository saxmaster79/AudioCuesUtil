using Microsoft.Extensions.DependencyInjection;
using System.CommandLine;

namespace AudioCuesUtil
{
    public static class CliCommandCollectionExtensions
    {
        public static IServiceCollection AddCliCommands(this IServiceCollection services)
        {
            Type greetCommandType = typeof(WriteToZipCommand);
            Type commandType = typeof(Command);

            IEnumerable<Type> commands = greetCommandType
                .Assembly
                .GetExportedTypes()
                .Where(x => x.Namespace == greetCommandType.Namespace && commandType.IsAssignableFrom(x) && !x.IsAbstract);

            foreach (Type command in commands)
            {
                services.AddSingleton(commandType, command);
            }

            return services;
        }
    }
}
