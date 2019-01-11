using System;
using System.Collections.Generic;
using System.Text;

namespace System.Resources
{
    internal static class MultitargetUtil
    {
        /// <devdoc>
        ///     This method gets assembly info for the corresponding type. If the delegate
        ///     is provided it is used to get this information.
        /// </devdoc>
        public static string GetAssemblyQualifiedName(Type type, Func<Type, string> typeNameConverter)
        {
            string assemblyQualifiedName = null;

            if (type != null)
            {
                if (typeNameConverter != null)
                {
                    try
                    {
                        assemblyQualifiedName = typeNameConverter(type);
                    }
                    catch (Exception e)
                    {
                        if (ClientUtils.IsSecurityOrCriticalException(e))
                        {
                            throw;
                        }
                    }
                }

                if (string.IsNullOrEmpty(assemblyQualifiedName))
                {
                    assemblyQualifiedName = type.AssemblyQualifiedName;
                }
            }

            return assemblyQualifiedName;
        }
    }

    internal static class ClientUtils
    {
        public static bool IsCriticalException(Exception ex) {
            return ex is NullReferenceException
                    || ex is StackOverflowException
                    || ex is OutOfMemoryException
                    || ex is System.Threading.ThreadAbortException
                    || ex is IndexOutOfRangeException
                    || ex is AccessViolationException;
        }

        public static bool IsSecurityOrCriticalException(Exception ex) {
            return (ex is System.Security.SecurityException) || IsCriticalException(ex);
        }
    }
}
