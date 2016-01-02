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
using System.Text.RegularExpressions;

namespace JsonKit {
    internal class __Version {
        private static String _version;

        public static String Version {
            get { return _version ?? (_version = GetVersion()); }
        }

        private static String GetVersion() {
            return _version =
                "JsonKit "
                + Regex.Match(Assembly.GetExecutingAssembly().FullName, @"\d+.\d+.\d+.\d+").Value
                + Environment.NewLine
                + "Copyright (c) 2015 Roy Xu";
        }
    }
}
