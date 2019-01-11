using System;
using System.Collections.Generic;
using System.Text;

namespace System.Resources
{
    internal static class MultitargetUtil
    {
        internal static string GetAssemblyQualifiedName(Type type, Func<Type, string> func)
        {
            return "FIXME";
        }
    }

    internal static class ClientUtils
    {
        internal static bool IsCriticalException(Exception ex)
        {
            return false; // FIXME
        }

        internal static bool IsSecurityOrCriticalException(Exception ex)
        {
            return false; // FIXME
        }
    }
}
