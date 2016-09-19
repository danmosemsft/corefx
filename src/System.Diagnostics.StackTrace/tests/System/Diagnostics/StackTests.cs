// Licensed to the .NET Foundation under one or more agreements.
// The .NET Foundation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Reflection;
using Xunit;

namespace System.Diagnostics.Tests
{
    public static class StackTests
    {
        [Fact]
        public static void Basic()
        {
            StackTrace st = new StackTrace();

            StackFrame sf = st.GetFrame(0);

            Assert.Equal(MethodBase.GetCurrentMethod(), sf.GetMethod());
            Assert.Equal(null, sf.GetMethod());
        }
    }
}
