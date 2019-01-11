// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.
// ------------------------------------------------------------------------------
// Changes to this file must follow the http://aka.ms/api-review process.
// ------------------------------------------------------------------------------


namespace System.Resources
{
    public partial interface IResourceWriter : System.IDisposable
    {
        void AddResource(string name, byte[] value);
        void AddResource(string name, object value);
        void AddResource(string name, string value);
        void Close();
        void Generate();
    }
    public sealed partial class ResourceWriter : System.IDisposable, System.Resources.IResourceWriter
    {
        public ResourceWriter(System.IO.Stream stream) { }
        public ResourceWriter(string fileName) { }
        public System.Func<System.Type, string> TypeNameConverter { [System.Runtime.CompilerServices.CompilerGeneratedAttribute]get { throw null; } [System.Runtime.CompilerServices.CompilerGeneratedAttribute]set { } }
        public void AddResource(string name, byte[] value) { }
        public void AddResource(string name, System.IO.Stream value) { }
        public void AddResource(string name, System.IO.Stream value, bool closeAfterWrite) { }
        public void AddResource(string name, object value) { }
        public void AddResource(string name, string value) { }
        public void AddResourceData(string name, string typeName, byte[] serializedData) { }
        public void Close() { }
        public void Dispose() { }
        public void Generate() { }
    }
    public sealed partial class ResXDataNode : System.Runtime.Serialization.ISerializable
    {
        public ResXDataNode(string name, object value) { }
        public ResXDataNode(string name, object value, System.Func<System.Type, string> typeNameConverter) { }
        public ResXDataNode(string name, System.Resources.ResXFileRef fileRef) { }
        public ResXDataNode(string name, System.Resources.ResXFileRef fileRef, System.Func<System.Type, string> typeNameConverter) { }
        public string Comment { get { throw null; } set { } }
        public System.Resources.ResXFileRef FileRef { get { throw null; } }
        public string Name { get { throw null; } set { } }
        public System.Collections.Generic.KeyValuePair<int, int> GetNodePosition() { throw null; }
        public object GetValue(System.ComponentModel.Design.ITypeResolutionService typeResolver) { throw null; }
        public object GetValue(System.Reflection.AssemblyName[] names) { throw null; }
        public string GetValueTypeName(System.ComponentModel.Design.ITypeResolutionService typeResolver) { throw null; }
        public string GetValueTypeName(System.Reflection.AssemblyName[] names) { throw null; }
        void System.Runtime.Serialization.ISerializable.GetObjectData(System.Runtime.Serialization.SerializationInfo si, System.Runtime.Serialization.StreamingContext context) { }
    }
    [System.ComponentModel.TypeConverterAttribute(typeof(System.Resources.ResXFileRef.Converter))]
    public partial class ResXFileRef
    {
        public ResXFileRef(string fileName, string typeName) { }
        public ResXFileRef(string fileName, string typeName, System.Text.Encoding textFileEncoding) { }
        public string FileName { get { throw null; } }
        public System.Text.Encoding TextFileEncoding { get { throw null; } }
        public string TypeName { get { throw null; } }
        public override string ToString() { throw null; }
        public partial class Converter : System.ComponentModel.TypeConverter
        {
            public Converter() { }
            public override bool CanConvertFrom(System.ComponentModel.ITypeDescriptorContext context, System.Type sourceType) { throw null; }
            public override bool CanConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Type destinationType) { throw null; }
            public override object ConvertFrom(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value) { throw null; }
            public override object ConvertTo(System.ComponentModel.ITypeDescriptorContext context, System.Globalization.CultureInfo culture, object value, System.Type destinationType) { throw null; }
        }
    }
    public partial class ResXResourceReader : System.Collections.IEnumerable, System.IDisposable, System.Resources.IResourceReader
    {
        public ResXResourceReader(System.IO.Stream stream) { }
        public ResXResourceReader(System.IO.Stream stream, System.ComponentModel.Design.ITypeResolutionService typeResolver) { }
        public ResXResourceReader(System.IO.Stream stream, System.Reflection.AssemblyName[] assemblyNames) { }
        public ResXResourceReader(System.IO.TextReader reader) { }
        public ResXResourceReader(System.IO.TextReader reader, System.ComponentModel.Design.ITypeResolutionService typeResolver) { }
        public ResXResourceReader(System.IO.TextReader reader, System.Reflection.AssemblyName[] assemblyNames) { }
        public ResXResourceReader(string fileName) { }
        public ResXResourceReader(string fileName, System.ComponentModel.Design.ITypeResolutionService typeResolver) { }
        public ResXResourceReader(string fileName, System.Reflection.AssemblyName[] assemblyNames) { }
        public string BasePath { get { throw null; } set { } }
        public bool UseResXDataNodes { get { throw null; } set { } }
        public void Close() { }
        protected virtual void Dispose(bool disposing) { }
        ~ResXResourceReader() { }
        public static System.Resources.ResXResourceReader FromFileContents(string fileContents) { throw null; }
        public static System.Resources.ResXResourceReader FromFileContents(string fileContents, System.ComponentModel.Design.ITypeResolutionService typeResolver) { throw null; }
        public static System.Resources.ResXResourceReader FromFileContents(string fileContents, System.Reflection.AssemblyName[] assemblyNames) { throw null; }
        public System.Collections.IDictionaryEnumerator GetEnumerator() { throw null; }
        public System.Collections.IDictionaryEnumerator GetMetadataEnumerator() { throw null; }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }
        void System.IDisposable.Dispose() { }
    }
    public partial class ResXResourceSet : System.Resources.ResourceSet
    {
        public ResXResourceSet(System.IO.Stream stream) { }
        public ResXResourceSet(string fileName) { }
        public override System.Type GetDefaultReader() { throw null; }
        public override System.Type GetDefaultWriter() { throw null; }
    }
    public partial class ResXResourceWriter : System.IDisposable, System.Resources.IResourceWriter
    {
        public static readonly string BinSerializedObjectMimeType;
        public static readonly string ByteArraySerializedObjectMimeType;
        public static readonly string DefaultSerializedObjectMimeType;
        public static readonly string ResMimeType;
        public static readonly string ResourceSchema;
        public static readonly string SoapSerializedObjectMimeType;
        public static readonly string Version;
        public ResXResourceWriter(System.IO.Stream stream) { }
        public ResXResourceWriter(System.IO.Stream stream, System.Func<System.Type, string> typeNameConverter) { }
        public ResXResourceWriter(System.IO.TextWriter textWriter) { }
        public ResXResourceWriter(System.IO.TextWriter textWriter, System.Func<System.Type, string> typeNameConverter) { }
        public ResXResourceWriter(string fileName) { }
        public ResXResourceWriter(string fileName, System.Func<System.Type, string> typeNameConverter) { }
        public string BasePath { get { throw null; } set { } }
        public virtual void AddAlias(string aliasName, System.Reflection.AssemblyName assemblyName) { }
        public void AddMetadata(string name, byte[] value) { }
        public void AddMetadata(string name, object value) { }
        public void AddMetadata(string name, string value) { }
        public void AddResource(System.Resources.ResXDataNode node) { }
        public void AddResource(string name, byte[] value) { }
        public void AddResource(string name, object value) { }
        public void AddResource(string name, string value) { }
        public void Close() { }
        public virtual void Dispose() { }
        protected virtual void Dispose(bool disposing) { }
        ~ResXResourceWriter() { }
        public void Generate() { }
    }
}
