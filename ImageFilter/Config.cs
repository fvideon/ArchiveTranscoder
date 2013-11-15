using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Globalization;
using System.IO;

namespace ImageFilter {
    class Config {
        /// <summary>
        /// Use standard windows temp path unless overridden in the app.config
        /// </summary>
        public static string TempPath = System.IO.Path.GetTempPath();

        /// <summary>
        /// The directory where the ImageMagick tool 'compare.exe' is found.
        /// </summary>
        public static string ImageMagickPath = @"C:\Program Files\ImageMagick";

        /// <summary>
        /// This static constructor reads for every non-constant static field in the class and checks for an app.config override.
        /// </summary>
        static Config() {
            Type myType = typeof(Config);
            FieldInfo[] fields = myType.GetFields(BindingFlags.Public | BindingFlags.Static);
            foreach (FieldInfo field in fields) {
                if (field.IsLiteral) // Is Constant?
                    continue; // We can't set it

                string fullName = field.Name;
                string configOverride = System.Configuration.ConfigurationManager.AppSettings[fullName];
                if (configOverride != null && configOverride != String.Empty) {
                    Type newType = field.FieldType;
                    object newValue = Convert.ChangeType(configOverride, newType, CultureInfo.InvariantCulture);
                    field.SetValue(null, newValue);
                }
            }
        }
    }
}
