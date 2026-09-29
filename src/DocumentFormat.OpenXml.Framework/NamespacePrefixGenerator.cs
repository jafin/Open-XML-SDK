// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Globalization;

namespace DocumentFormat.OpenXml
{
    internal static class NamespacePrefixGenerator
    {
        /// <summary>
        /// Generates a prefix of the form <c>ns0</c>, <c>ns1</c>, ... that <paramref name="isDeclared"/> reports as not declared.
        /// </summary>
        public static string Generate(Func<string, bool> isDeclared)
        {
            for (var i = 0; ; i++)
            {
                var prefix = "ns" + i.ToString(CultureInfo.InvariantCulture);

                if (!isDeclared(prefix))
                {
                    return prefix;
                }
            }
        }
    }
}
