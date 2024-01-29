using System.CommandLine;

namespace AudioCuesUtil;
public abstract class BaseCommand : Command
{
    protected BaseCommand(Type type, string description) : base(type.Name.Replace("Command", ""), description)
    {
    }
}