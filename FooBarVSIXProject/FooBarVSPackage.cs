//------------------------------------------------------------------------------
// <copyright file="FooBarVSPackage.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System.ComponentModel;
using System.Collections;
using System.Linq;
using PropertyGridSample;

namespace FooBarVSIXProject
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(FooBarVSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideOptionPage(typeof(OptionPageGrid),
    "My Category", "My Grid Page", 0, 0, true)]
    public sealed class FooBarVSPackage : Package
    {
        /// <summary>
        /// FooBarVSPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "38f26907-fa66-4a4a-a2fb-4604c8a8a1b7";

        /// <summary>
        /// Initializes a new instance of the <see cref="FooBarVSPackage"/> class.
        /// </summary>
        public FooBarVSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        protected override void Initialize()
        {
            base.Initialize();
        }

        #endregion
    }

    [Guid("AA78A733-6E10-45B9-ACA3-4DB94A16D151")]
    public class OptionPageGrid : DialogPage
    {
        [Category("General")]
        [DisplayName("Foos")]
        [Description("Bla Foo Bla")]
        [TypeConverter(typeof(StringArrayConverter))]
        public string[] Foos
        { get; set; }

        [Category("General")]
        [DisplayName("Bar")]
        [Description("Bla Bar Bla")]
        [TypeConverter(typeof(CustomIntConverter))]
        public int Bar
        { get; set; }

        [Category("General")]
        [DisplayName("Baz")]
        [Description("Bla Baz Bla")]
        [TypeConverter(typeof(IntArrayConverter))]
        public int[] Baz
        { get; set; }

        public override void SaveSettingsToStorage()
        {
            base.SaveSettingsToStorage();
        }
    }

    public class StringArrayConverter : TypeConverter
    {
        private const string delimiter = "#@#";

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return destinationType == typeof(string[]) || base.CanConvertTo(context, destinationType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            string v = value as string;

            return v == null ? base.ConvertFrom(context, culture, value) : v.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries);
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            string[] v = value as string[];
            if (destinationType != typeof(string) || v == null)
            {
                return base.ConvertTo(context, culture, value, destinationType);
            }
            return string.Join(delimiter, v);
        }
    }

    public class IntArrayConverter : TypeConverter
    {
        private const string delimiter = "#@#";

        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return destinationType == typeof(string);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            string v = value as string;

            return v == null ? base.ConvertFrom(context, culture, value) : v.Split(new[] { delimiter }, StringSplitOptions.RemoveEmptyEntries).Select(int.Parse).ToArray();
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            var v = value as int[];
            if (destinationType != typeof(string) || v == null)
            {
                return base.ConvertTo(context, culture, value, destinationType);
            }
            return string.Join(delimiter, v.Select(n => n.ToString()).ToArray());
        }
    }

    public class CustomIntConverter : TypeConverter
    {
        public override bool CanConvertFrom(ITypeDescriptorContext context, Type sourceType)
        {
            return sourceType == typeof(string) || base.CanConvertFrom(context, sourceType);
        }

        public override bool CanConvertTo(ITypeDescriptorContext context, Type destinationType)
        {
            return destinationType == typeof(int) || base.CanConvertTo(context, destinationType);
        }

        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            string v = value as string;

            return int.Parse(v.TrimStart('*'));
        }

        public override object ConvertTo(ITypeDescriptorContext context, CultureInfo culture, object value, Type destinationType)
        {
            var v = (int)value;
            if (destinationType != typeof(string))
            {
                return base.ConvertTo(context, culture, value, destinationType);
            }
            return v.ToString().PadLeft(25, '*');
        }
    }
}
