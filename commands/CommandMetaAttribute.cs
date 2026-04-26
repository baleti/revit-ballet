using System;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class CommandMetaAttribute : Attribute
{
    public string Input { get; }
    public CommandMetaAttribute(string input) { Input = input; }
}

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class CommandOutputAttribute : Attribute
{
    public string Output { get; }
    public CommandOutputAttribute(string output) { Output = output; }
}
