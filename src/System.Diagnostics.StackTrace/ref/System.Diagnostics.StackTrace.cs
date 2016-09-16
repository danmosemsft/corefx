// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.
// ------------------------------------------------------------------------------
// Changes to this file must follow the http://aka.ms/api-review process.
// ------------------------------------------------------------------------------


namespace System.Diagnostics
{
    public partial class StackFrame
    {
        public const int OFFSET_UNKNOWN = -1;
        public StackFrame() { }
        public StackFrame(bool fNeedFileInfo) { }
        public StackFrame(int skipFrames) { }
        public StackFrame(int skipFrames, bool fNeedFileInfo) { }
        public StackFrame(string fileName, int lineNumber) { }
        public StackFrame(string fileName, int lineNumber, int colNumber) { }
        public virtual int GetFileColumnNumber() { throw null; }
        public virtual int GetFileLineNumber() { throw null; }
        public virtual string GetFileName() { throw null; }
        public virtual int GetILOffset() { throw null; }
        public virtual System.Reflection.MethodBase GetMethod() { throw null; }
        public virtual int GetNativeOffset() { throw null; }
        public override string ToString() { throw null; }
    }
    public static partial class StackFrameExtensions
    {
        public static System.IntPtr GetNativeImageBase(this System.Diagnostics.StackFrame stackFrame) { return default(System.IntPtr); }
        public static System.IntPtr GetNativeIP(this System.Diagnostics.StackFrame stackFrame) { return default(System.IntPtr); }
        public static bool HasILOffset(this System.Diagnostics.StackFrame stackFrame) { return default(bool); }
        public static bool HasMethod(this System.Diagnostics.StackFrame stackFrame) { return default(bool); }
        public static bool HasNativeImage(this System.Diagnostics.StackFrame stackFrame) { return default(bool); }
        public static bool HasSource(this System.Diagnostics.StackFrame stackFrame) { return default(bool); }
    }
    public partial class StackTrace
    {
        public const int METHODS_TO_SKIP = 0;
        public StackTrace() { }
        public StackTrace(bool fNeedFileInfo) { }
        public StackTrace(System.Diagnostics.StackFrame frame) { }
        public StackTrace(System.Exception e) { }
        public StackTrace(System.Exception e, bool fNeedFileInfo) { }
        public StackTrace(System.Exception e, int skipFrames) { }
        public StackTrace(System.Exception e, int skipFrames, bool fNeedFileInfo) { }
        public StackTrace(int skipFrames) { }
        public StackTrace(int skipFrames, bool fNeedFileInfo) { }
        [System.ObsoleteAttribute("This constructor has been deprecated.  Please use a constructor that does not require a Thread parameter.  http://go.microsoft.com/fwlink/?linkid=14202")]
        public StackTrace(System.Threading.Thread targetThread, bool needFileInfo) { }
        public virtual int FrameCount { get { throw null; } }
        public virtual System.Diagnostics.StackFrame GetFrame(int index) { throw null; }
        public virtual System.Diagnostics.StackFrame[] GetFrames() { throw null; }
        public override string ToString() { throw null; }
    }
}
