/*
 * The MIT License (MIT)
 *
 * Copyright (c) 2015 Roy Xu
 *
 * @since 12/31/2015 18:53:05
 *
 * @author Roy Xu
*/

using System;
using System.Reflection;
using System.Runtime.Versioning;

namespace kitapp
{
    internal class __Version
    {
        private static String _version;

        public static String Version
        {
            get { return _version ?? (_version = GetVersion()); }
        }

        private static String GetVersion()
        {
            var executingAssembly = Assembly.GetExecutingAssembly();
            var title = executingAssembly.GetCustomAttribute<AssemblyTitleAttribute>().Title;
            var copyright = executingAssembly.GetCustomAttribute<AssemblyCopyrightAttribute>().Copyright;
            var version = executingAssembly.GetCustomAttribute<AssemblyFileVersionAttribute>().Version;

            return _version = $"{title} v{version} {copyright}";
        }
    }
}
