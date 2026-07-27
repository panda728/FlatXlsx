// The trimming and AOT annotations live in the framework from .NET 7 onward. Declaring them
// here for the older targets lets the entry points carry the attribute unconditionally, instead
// of wrapping every one of them in #if. On those targets the attributes are inert metadata -
// no analyzer reads them - which is correct, because trimming and native AOT are features of
// the modern runtimes only.
#if !NET8_0_OR_GREATER
namespace System.Diagnostics.CodeAnalysis
{
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class,
        Inherited = false)]
    internal sealed class RequiresUnreferencedCodeAttribute(string message) : Attribute
    {
        public string Message { get; } = message;
        public string? Url { get; set; }
    }

    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Constructor | AttributeTargets.Class,
        Inherited = false)]
    internal sealed class RequiresDynamicCodeAttribute(string message) : Attribute
    {
        public string Message { get; } = message;
        public string? Url { get; set; }
    }
}
#endif
