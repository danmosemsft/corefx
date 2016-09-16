// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.
// ------------------------------------------------------------------------------
// Changes to this file must follow the http://aka.ms/api-review process.
// ------------------------------------------------------------------------------


namespace System.Diagnostics
{
    public static partial class Debug
    {
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Assert(bool condition) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Assert(bool condition, string message) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Assert(bool condition, string message, string detailMessage) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Assert(bool condition, string message, string detailMessageFormat, params object[] args) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Fail(string message) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Fail(string message, string detailMessage) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Write(object value) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Write(object value, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Write(string message) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void Write(string message, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteIf(bool condition, object value) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteIf(bool condition, object value, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteIf(bool condition, string message) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteIf(bool condition, string message, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLine(object value) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLine(object value, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLine(string message) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLine(string format, params object[] args) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLine(string message, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLineIf(bool condition, object value) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLineIf(bool condition, object value, string category) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLineIf(bool condition, string message) { }
        [System.Diagnostics.ConditionalAttribute("DEBUG")]
        public static void WriteLineIf(bool condition, string message, string category) { }
    }  
    public sealed partial class Debugger
    {
        public static readonly string DefaultCategory;
        [System.ObsoleteAttribute("Do not create instances of the Debugger class.  Call the static methods directly on this type instead", true)]
        public Debugger() { }
        public static bool IsAttached { get { throw null; } }
        public static void Break() { }
        public static bool IsLogging() { throw null; }
        public static bool Launch() { throw null; }
        public static void Log(int level, string category, string message) { }
        public static void NotifyOfCrossThreadDependency() { }
    }
    [System.AttributeUsageAttribute((System.AttributeTargets)(384), AllowMultiple = false)]
    public sealed partial class DebuggerBrowsableAttribute : System.Attribute
    {
        public DebuggerBrowsableAttribute(System.Diagnostics.DebuggerBrowsableState state) { }
        public System.Diagnostics.DebuggerBrowsableState State { get { return default(System.Diagnostics.DebuggerBrowsableState); } }
    }
    public enum DebuggerBrowsableState
    {
        Collapsed = 2,
        Never = 0,
        RootHidden = 3,
    }
    [System.AttributeUsageAttribute((System.AttributeTargets)(4509), AllowMultiple = true)]
    public sealed partial class DebuggerDisplayAttribute : System.Attribute
    {
        public DebuggerDisplayAttribute(string value) { }
        public string Name { get { return default(string); } set { } }
        public System.Type Target { get { return default(System.Type); } set { } }
        public string TargetTypeName { get { return default(string); } set { } }
        public string Type { get { return default(string); } set { } }
        public string Value { get { return default(string); } }
    }
    [System.AttributeUsageAttribute((System.AttributeTargets)(224), Inherited = false)]
    public sealed partial class DebuggerHiddenAttribute : System.Attribute
    {
        public DebuggerHiddenAttribute() { }
    }
    [System.AttributeUsageAttribute((System.AttributeTargets)(236), Inherited = false)]
    public sealed partial class DebuggerNonUserCodeAttribute : System.Attribute
    {
        public DebuggerNonUserCodeAttribute() { }
    }
    [System.AttributeUsageAttribute((System.AttributeTargets)(108), Inherited = false)]
    public sealed partial class DebuggerStepThroughAttribute : System.Attribute
    {
        public DebuggerStepThroughAttribute() { }
    }
    [System.AttributeUsageAttribute((System.AttributeTargets)(13), AllowMultiple = true)]
    public sealed partial class DebuggerTypeProxyAttribute : System.Attribute
    {
        public DebuggerTypeProxyAttribute(string typeName) { }
        public DebuggerTypeProxyAttribute(System.Type type) { }
        public string ProxyTypeName { get { return default(string); } }
        public System.Type Target { get { return default(System.Type); } set { } }
        public string TargetTypeName { get { return default(string); } set { } }
    }
}
