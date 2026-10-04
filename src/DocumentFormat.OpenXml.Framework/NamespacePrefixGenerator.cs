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
                var prefix = Prefix + i.ToString(CultureInfo.InvariantCulture);

                if (!isDeclared(prefix))
                {
                    return prefix;
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether <paramref name="prefix"/> has the form of a generated prefix, so it may clash with one.
        /// </summary>
        public static bool CouldGenerate(string prefix)
        {
            if (prefix.Length <= Prefix.Length || !prefix.StartsWith(Prefix, StringComparison.Ordinal))
            {
                return false;
            }

            for (var i = Prefix.Length; i < prefix.Length; i++)
            {
                if (prefix[i] is < '0' or > '9')
                {
                    return false;
                }
            }

            return true;
        }

        private const string Prefix = "ns";
    }
}
