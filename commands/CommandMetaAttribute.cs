using System;

[AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
public class CommandMetaAttribute : Attribute
{
    public string Input { get; }
    public CommandMetaAttribute(string input) { Input = input; }
}
