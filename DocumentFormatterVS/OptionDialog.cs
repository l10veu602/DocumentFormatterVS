using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;
using Microsoft.VisualStudio.Shell.Settings;
using Microsoft.VisualStudio.Settings;

namespace DocumentFormatterVS
{
    class OptionDialog : DialogPage
    {
        class PathRegexesTypeConverter : TypeConverter
        {
            private const string Delimeter = "::";

            public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
            {
                return sourceType == typeof(string);
            }

            public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
            {
                string serialized = value as string;

                if (serialized == null)
                {
                    return base.ConvertFrom(context, culture, value);
                }
                else if (serialized == string.Empty)
                {
                    return new string[0];
                }

                return serialized.Split(new string[] { Delimeter }, StringSplitOptions.None);
            }

            public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
            {
                return destinationType == typeof(string);
            }

            public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
            {
                if (destinationType != typeof(string))
                {
                    return base.ConvertTo(context, culture, value, destinationType);
                }

                string[] pathRegexes = value as string[];

                if (pathRegexes == null)
                {
                    return string.Empty;
                }

                return string.Join(Delimeter, pathRegexes);
            }
        }

        private static string NameOfIgnorePathRegexes = nameof(IgnorePathRegexes);

        [Category("Ignore")]
        [DisplayName("Ignore list regex")]
        [Description("Ignore list regex")]
        public string[] IgnorePathRegexes
        {
            get; set;
        }

        [DisplayName("Is disabled")]
        [Description("Whether formatting is disabled")]
        public bool IsDisabled
        {
            get; set;
        }

        private string CollectionPath { get { return SharedSettingsStorePath.Replace('.', '\\'); } }

        protected override void LoadSettingFromStorage(PropertyDescriptor prop)
        {
            if (prop.Name == NameOfIgnorePathRegexes)
            {
                var userSettingsStore = GetWritableSettingStore();
                if (!userSettingsStore.PropertyExists(CollectionPath, NameOfIgnorePathRegexes))
                    return;

                string value = userSettingsStore.GetString(CollectionPath, NameOfIgnorePathRegexes);
                var converter = new PathRegexesTypeConverter();
                IgnorePathRegexes = converter.ConvertFrom(value) as string[];
            }
            else
            {
                base.LoadSettingFromStorage(prop);
            }
        }

        protected override void SaveSetting(PropertyDescriptor property)
        {
            if (property.Name == NameOfIgnorePathRegexes)
            {
                var converter = new PathRegexesTypeConverter();
                string value = converter.ConvertToInvariantString(IgnorePathRegexes);
                var userSettingsStore = GetWritableSettingStore();
                WriteSetting(userSettingsStore, CollectionPath, NameOfIgnorePathRegexes, value);
            }
            else
            {
                base.SaveSetting(property);
            }
        }

        private WritableSettingsStore GetWritableSettingStore()
        {
            var settingsManager = new ShellSettingsManager(Site);
            var userSettingsStore = settingsManager.GetWritableSettingsStore(SettingsScope.UserSettings);

            return userSettingsStore;
        }

        private void WriteSetting(WritableSettingsStore settingStore, string collectionName, string name, string value)
        {
            if (!settingStore.CollectionExists(collectionName))
                settingStore.CreateCollection(collectionName);

            settingStore.SetString(collectionName, name, value);
        }
    }
}
