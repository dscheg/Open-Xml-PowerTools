// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using OpenXmlPowerTools.HtmlToWml;
using OpenXmlPowerTools.HtmlToWml.CSS;
using System.Text.RegularExpressions;

namespace OpenXmlPowerTools
{
    public class HtmlToWmlConverterSettings
    {
        public string MajorLatinFont;
        public string MinorLatinFont;
        public double DefaultFontSize;
        public XElement DefaultSpacingElement;
        public XElement DefaultSpacingElementForParagraphsInTables;
        public XElement SectPr;
        public string DefaultBlockContentMargin;
        public string BaseUriForImages;

        public Twip PageWidthTwips { get { return (long)SectPr.Elements(W.pgSz).Attributes(W._w).FirstOrDefault(); } }
        public Twip PageMarginLeftTwips { get { return (long)SectPr.Elements(W.pgMar).Attributes(W.left).FirstOrDefault(); } }
        public Twip PageMarginRightTwips { get { return (long)SectPr.Elements(W.pgMar).Attributes(W.right).FirstOrDefault(); } }
        public Emu PageWidthEmus { get { return Emu.TwipsToEmus(PageWidthTwips); } }
        public Emu PageMarginLeftEmus { get { return Emu.TwipsToEmus(PageMarginLeftTwips); } }
        public Emu PageMarginRightEmus { get { return Emu.TwipsToEmus(PageMarginRightTwips); } }
    }

    public class HtmlToWmlConverter
    {
        public static WmlDocument ConvertHtmlToWml(
            string defaultCss,
            string authorCss,
            string userCss,
            XElement xhtml,
            HtmlToWmlConverterSettings settings)
        {
            return HtmlToWmlConverterCore.ConvertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings, null, null);
        }

        public static WmlDocument ConvertHtmlToWml(
            string defaultCss,
            string authorCss,
            string userCss,
            XElement xhtml,
            HtmlToWmlConverterSettings settings,
            WmlDocument emptyDocument,
            string annotatedHtmlDumpFileName)
        {
            return HtmlToWmlConverterCore.ConvertHtmlToWml(defaultCss, authorCss, userCss, xhtml, settings, emptyDocument, annotatedHtmlDumpFileName);
        }

        private static string s_Blank_wml_base64 = @"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";

        private static WmlDocument s_EmptyDocument = null;

        public static WmlDocument EmptyDocument
        {
            get {
                if (s_EmptyDocument == null)
                {
                    s_EmptyDocument = new WmlDocument("EmptyDocument.docx", Convert.FromBase64String(s_Blank_wml_base64));
                }
                return s_EmptyDocument;
            }
        }

        public static HtmlToWmlConverterSettings GetDefaultSettings()
        {
            return GetDefaultSettings(EmptyDocument);
        }

        public static HtmlToWmlConverterSettings GetDefaultSettings(WmlDocument wmlDocument)
        {
            HtmlToWmlConverterSettings settings = new HtmlToWmlConverterSettings();
            using (MemoryStream ms = new MemoryStream())
            {
                ms.Write(wmlDocument.DocumentByteArray, 0, wmlDocument.DocumentByteArray.Length);
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(ms, false))
                {
                    string majorLatinFont, minorLatinFont;
                    double defaultFontSize;
                    GetDefaultFontInfo(wDoc, out majorLatinFont, out minorLatinFont, out defaultFontSize);
                    settings.MajorLatinFont = majorLatinFont;
                    settings.MinorLatinFont = minorLatinFont;
                    settings.DefaultFontSize = defaultFontSize;

                    settings.MinorLatinFont = "Times New Roman";
                    settings.DefaultFontSize = 12d;
                    settings.DefaultBlockContentMargin = "auto";
                    settings.DefaultSpacingElement = new XElement(W.spacing,
                        new XAttribute(W.before, 100),
                        new XAttribute(W.beforeAutospacing, 1),
                        new XAttribute(W.after, 100),
                        new XAttribute(W.afterAutospacing, 1),
                        new XAttribute(W.line, 240),
                        new XAttribute(W.lineRule, "auto"));
                    settings.DefaultSpacingElementForParagraphsInTables = new XElement(W.spacing,
                        new XAttribute(W.before, 100),
                        new XAttribute(W.beforeAutospacing, 1),
                        new XAttribute(W.after, 100),
                        new XAttribute(W.afterAutospacing, 1),
                        new XAttribute(W.line, 240),
                        new XAttribute(W.lineRule, "auto"));

                    XDocument mXDoc = wDoc.MainDocumentPart.GetXDocument();
                    XElement existingSectPr = mXDoc.Root.Descendants(W.sectPr).FirstOrDefault();
                    settings.SectPr = new XElement(W.sectPr,
                        existingSectPr.Elements(W.pgSz),
                        existingSectPr.Elements(W.pgMar));
                }
            }
            return settings;
        }

        private static void GetDefaultFontInfo(WordprocessingDocument wDoc, out string majorLatinFont, out string minorLatinFont, out double defaultFontSize)
        {
            if (wDoc.MainDocumentPart.ThemePart != null)
            {
                XElement fontScheme = wDoc.MainDocumentPart.ThemePart.GetXDocument().Root.Elements(A.themeElements).Elements(A.fontScheme).FirstOrDefault();
                if (fontScheme != null)
                {
                    majorLatinFont = (string)fontScheme.Elements(A.majorFont).Elements(A.latin).Attributes(NoNamespace.typeface).FirstOrDefault();
                    minorLatinFont = (string)fontScheme.Elements(A.minorFont).Elements(A.latin).Attributes(NoNamespace.typeface).FirstOrDefault();
                    string defaultFontSizeString = (string)wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument().Root.Elements(W.docDefaults)
                        .Elements(W.rPrDefault).Elements(W.rPr).Elements(W.sz).Attributes(W.val).FirstOrDefault();
                    if (defaultFontSizeString != null)
                    {
                        double dfs;
                        if (double.TryParse(defaultFontSizeString, out dfs))
                        {
                            defaultFontSize = dfs / 2d;
                            return;
                        }
                        defaultFontSize = 12;
                        return;
                    }
                }
            }
            majorLatinFont = "";
            minorLatinFont = "";
            defaultFontSize = 12;
        }

        public static string CleanUpCss(string css)
        {
            if (css == null)
                return "";
            css = css.Trim();
            string cleanCss = Regex.Split(css, "\r\n|\r|\n")
                .Where(l =>
                {
                    string lTrim = l.Trim();
                    if (lTrim == "//")
                        return false;
                    if (lTrim == "////")
                        return false;
                    if (lTrim == "<!--" || lTrim == "&lt;!--")
                        return false;
                    if (lTrim == "-->" || lTrim == "--&gt;")
                        return false;
                    return true;
                })
                .Select(l => l + Environment.NewLine )
                .StringConcatenate();
            return cleanCss;
        }
    }

    public struct Emu
    {
        public long m_Value;
        public static int s_EmusPerInch = 914400;

        public static Emu TwipsToEmus(long twips)
        {
            float v1 = (float)twips / 20f;
            float v2 = v1 / 72f;
            float v3 = v2 * s_EmusPerInch;
            long emus = (long)v3;
            return new Emu(emus);
        }

        public static Emu PointsToEmus(double points)
        {
            double v1 = points / 72;
            double v2 = v1 * s_EmusPerInch;
            long emus = (long)v2;
            return new Emu(emus);
        }

        public Emu(long value)
        {
            m_Value = value;
        }

        public static implicit operator long(Emu e)
        {
            return e.m_Value;
        }

        public static implicit operator Emu(long l)
        {
            return new Emu(l);
        }

        public override string ToString()
        {
            throw new OpenXmlPowerToolsException("Can't convert directly to string, must cast to long");
        }
    }

    public struct TPoint
    {
        public double m_Value;

        public TPoint(double value)
        {
            m_Value = value;
        }

        public static implicit operator double(TPoint t)
        {
            return t.m_Value;
        }

        public static implicit operator TPoint(double d)
        {
            return new TPoint(d);
        }

        public override string ToString()
        {
            throw new OpenXmlPowerToolsException("Can't convert directly to string, must cast to double");
        }
    }

    public struct Twip
    {
        public long m_Value;

        public Twip(long value)
        {
            m_Value = value;
        }

        public static implicit operator long(Twip t)
        {
            return t.m_Value;
        }

        public static implicit operator Twip(long l)
        {
            return new Twip(l);
        }

        public static implicit operator Twip(double d)
        {
            return new Twip((long)d);
        }

        public override string ToString()
        {
            throw new OpenXmlPowerToolsException("Can't convert directly to string, must cast to long");
        }
    }

    public class SizeEmu
    {
        public Emu m_Height;
        public Emu m_Width;

        public SizeEmu(Emu width, Emu height)
        {
            m_Width = width;
            m_Height = height;
        }

        public SizeEmu(long width, long height)
        {
            m_Width = width;
            m_Height = height;
        }
    }
}

