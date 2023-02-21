using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using FelisOpenXml.FelisShape.Base;
using FelisOpenXml.FelisShape.Draw;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape.Text
{
    /// <summary>
    /// The class of the properties of the text
    /// </summary>
    public class FelisTextProperties : FelisCompositeElement
    {
        /// <summary>
        /// The parent element
        /// </summary>
        protected readonly OpenXmlElement? ParentElement;

        /// <summary>
        /// Create an instance assigned to a special text's properties
        /// </summary>
        /// <param name="_element">The element of the special text's properties</param>
        /// <param name="_parentElement">The parant the text's properties belonged to.</param>
        internal FelisTextProperties(A.TextCharacterPropertiesType _element, OpenXmlElement? _parentElement)
            : base(_element)
        {
            ParentElement = _parentElement;
        }

        /// <summary>
        /// Submit the changing in the properties. 
        /// This method is only used when the properties object is created in runing time.
        /// The changing will be submitted immediately if the properties is already existed in the DOM tree of the parent element.
        /// </summary>
        /// <returns>True for success. False for fail. If the properties is not assigned to a parent the method will be fail.</returns>
        public bool Submit()
        {
            try
            {
                if (null != ParentElement)
                {
                    if (!ParentElement.Contains(Element))
                    {
                        if (Element is A.EndParagraphRunProperties)
                        {
                            ParentElement.InsertElement(Element);
                        }
                        else
                        {
                            ParentElement.InsertElement(Element, 0);
                        }
                    }

                    return true;
                }
            }
            catch (Exception err)
            {
                Trace.TraceWarning(err.ToString());
            }

            return false;
        }

        /// <summary>
        /// Copy the properies from a given object
        /// </summary>
        /// <param name="_sourceProps">The source object containing the properties should be copied.</param>
        /// <param name="_isPure">Set true for clearing all existed defined properties in the target object before copy.</param>
        public void CopyFrom(FelisTextProperties? _sourceProps, bool _isPure)
        {
            if (null != _sourceProps)
            {
                if (_isPure)
                {
                    Element.ClearAllAttributes();
                    Element.RemoveAllChildren();
                }

                foreach (var srcChild in _sourceProps.Element.Elements())
                {
                    var existedChildren = Element.Elements().Where(e => e.GetType() == srcChild.GetType()).ToArray();
                    if (existedChildren.Length > 0)
                    {
                        Element.InsertBefore(srcChild.CloneNode(true), existedChildren[0]);
                        foreach (var existedChild in existedChildren)
                        {
                            existedChild.Remove();
                        }
                    }
                    else
                    {
                        Element.Append(srcChild.CloneNode(true));
                    }
                }

                Element.SetAttributes(_sourceProps.Element.GetAttributes());
            }
        }

        /// <summary>
        /// The size property of the font
        /// </summary>
        public int? FontSize
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.FontSize?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.FontSize = value;
                }
            }
        }

        /// <summary>
        /// The language property
        /// </summary>
        public string? Language
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Language?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.Language = value;
                }
            }
        }

        /// <summary>
        /// The alternative language property
        /// </summary>
        public string? AlternativeLanguage
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.AlternativeLanguage?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.AlternativeLanguage = value;
                }
            }
        }

        /// <summary>
        /// The bold property
        /// </summary>
        public bool? Bold
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Bold?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.Bold = value;
                }
            }
        }

        /// <summary>
        /// The italic  property
        /// </summary>
        public bool? Italic
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Italic?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.Italic = value;
                }
            }
        }

        /// <summary>
        /// The underline  property
        /// </summary>
        public string? Underline
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Underline?.Value.ToString();
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    if (Enum.TryParse(typeof(A.TextUnderlineValues), value, true, out object? newVal))
                    {
                        if (newVal is A.TextUnderlineValues ulVal)
                        {
                            props.Underline = ulVal;
                            return;
                        }
                    }
                    props.Underline = null;
                }
            }
        }

        /// <summary>
        /// The baseline property
        /// </summary>
        public int? Baseline
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Baseline?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.Baseline = value;
                }
            }
        }

        /// <summary>
        /// The capital style property
        /// </summary>
        public string? Capital
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Capital?.Value.ToString();
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    if (Enum.TryParse(typeof(A.TextCapsValues), value, true, out object? newVal))
                    {
                        if (newVal is A.TextCapsValues tcVal)
                        {
                            props.Capital = tcVal;
                            return;
                        }
                    }
                    props.Capital = null;
                }
            }
        }

        /// <summary>
        /// The baseline property
        /// </summary>
        public int? Kerning
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Kerning?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.Kerning = value;
                }
            }
        }

        /// <summary>
        /// The normalize height property
        /// </summary>
        public bool? NormalizeHeight
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.NormalizeHeight?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.NormalizeHeight = value;
                }
            }
        }

        /// <summary>
        /// The spacing property
        /// </summary>
        public int? Spacing
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Spacing?.Value;
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    props.Spacing = value;
                }
            }
        }

        /// <summary>
        /// The strike style property
        /// </summary>
        public string? Strike
        {
            get
            {
                return (Element as A.TextCharacterPropertiesType)?.Strike?.Value.ToString();
            }

            set
            {
                if (Element is A.TextCharacterPropertiesType props)
                {
                    if (Enum.TryParse(typeof(A.TextStrikeValues), value, true, out object? newVal))
                    {
                        if (newVal is A.TextStrikeValues tsVal)
                        {
                            props.Strike = tsVal;
                            return;
                        }
                    }
                    props.Strike = null;
                }
            }
        }

        internal static readonly IReadOnlyDictionary<Type, string> FontType2Names = new Dictionary<Type, string>()
        {
            { typeof(A.LatinFont), "Latin" },
            { typeof(A.EastAsianFont), "EastAsian" },
            { typeof(A.SymbolFont), "Symbol" },
            { typeof(A.ComplexScriptFont), "ComplexScript" },
            { typeof(A.BulletFont), "Bullet" }
        };

        internal static readonly IReadOnlyDictionary<string, Type> Name2FontType = (new Func<IReadOnlyDictionary<string, Type>>(() =>
        {
            var dict = new Dictionary<string, Type>();
            foreach (var item in FontType2Names)
            {
                dict[item.Value] = item.Key;
            }
            return dict;
        }))();

        /// <summary>
        /// Get an iterator of all the font families' names
        /// </summary>
        public IEnumerable<KeyValuePair<string, string>> FontFamilies
        {
            get
            {
                foreach (var fontTypeElement in Element.Elements<A.TextFontType>())
                {
                    if (FontType2Names.TryGetValue(fontTypeElement.GetType(), out string? _type) && !string.IsNullOrWhiteSpace(_type))
                    {
                        var family = GetNormalizeFontFamily(fontTypeElement);
                        if (null != family)
                        {
                            yield return new KeyValuePair<string, string>(_type, family);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Set the special font type to use the special font family.
        /// The latin and east asian font will be set if there is not any font type specialized.
        /// </summary>
        /// <param name="_familyName">The special font family</param>
        /// <param name="_fontTypes">The list of the font types wanna be set.</param>
        public void SetFontFamily(string _familyName, params string[] _fontTypes)
        {
            if (!string.IsNullOrWhiteSpace(_familyName))
            {
                var emptyTypes = Array.Empty<Type>();
                Type?[] types = _fontTypes.SelectMany(e => Name2FontType.TryGetValue(e, out Type? type) ? new[] { type } : emptyTypes).ToArray();
                if ((null == types) || (types.Length <= 0))
                {
                    types = new[] { typeof(A.LatinFont), typeof(A.EastAsianFont) };
                }
                foreach (var font in Element.Elements<A.TextFontType>())
                {
                    var idx = Array.IndexOf(types, font.GetType());
                    if (idx >= 0)
                    {
                        types[idx] = null;
                        font.Typeface = _familyName;
                    }
                }
                foreach (var type in types)
                {
                    if (null != type)
                    {
                        var newFont = Activator.CreateInstance(type) as A.TextFontType;
                        if (null != newFont)
                        {
                            newFont.Typeface = _familyName;
                            Element.AddChild(newFont, false);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Set the special font type to keep the same as the theme.
        /// The latin and east asian font will be set if there is not any font type specialized.
        /// </summary>
        /// <param name="_asMajor">True to set as the major text, False to set as the minor text.</param>
        /// <param name="_fontTypes">The list of the font types wanna be set.</param>
        public void SetFontFamilyAsTheme(bool _asMajor, params string[] _fontTypes)
        {
            var emptyTypes = Array.Empty<Type>();
            Type?[] types = _fontTypes.SelectMany(e => Name2FontType.TryGetValue(e, out Type? type) ? new[] { type } : emptyTypes).ToArray();
            if ((null == types) || (types.Length <= 0))
            {
                types = new[] { typeof(A.LatinFont), typeof(A.EastAsianFont) };
            }
            foreach (var font in Element.Elements<A.TextFontType>())
            {
                var idx = Array.IndexOf(types, font.GetType());
                if (idx >= 0)
                {
                    types[idx] = null;
                    font.Typeface = $"+{(_asMajor ? "mj" : "mn")}-{GetFontTypeAbbreviation(font)}";
                }
            }
            foreach (var type in types)
            {
                if (null != type)
                {
                    A.TextFontType? newFont = (type == typeof(A.LatinFont))
                                                ? new A.LatinFont()
                                                : ((type == typeof(A.ComplexScriptFont))
                                                    ? new A.ComplexScriptFont()
                                                    : ((type == typeof(A.EastAsianFont))
                                                        ? new A.EastAsianFont()
                                                        : ((type == typeof(A.SymbolFont))
                                                            ? new A.SymbolFont()
                                                            : ((type == typeof(A.BulletFont))
                                                                ? new A.BulletFont() : null))));
                    if (null != newFont)
                    {
                        newFont.Typeface = $"+{(_asMajor ? "mj" : "mn")}-{GetFontTypeAbbreviation(newFont)}";
                        Element.AddChild(newFont, false);
                    }
                }
            }
        }

        /// <summary>
        /// Get the font family used by the special font type
        /// </summary>
        /// <param name="_fontType">The special font type</param>
        /// <returns></returns>
        public string? GetFontFamily(string _fontType)
        {
            if (Name2FontType.TryGetValue(_fontType, out Type? type) && (null != type))
            {
                var font = Element.Elements<A.TextFontType>().FirstOrDefault(e => type.IsInstanceOfType(e));
                return GetNormalizeFontFamily(font);
            }
            return null;
        }

        /// <summary>
        /// Remove the special font type declaretion
        /// </summary>
        /// <param name="_fontType">The special font type</param>
        public void RemoveFontFamily(string _fontType)
        {
            if (Name2FontType.TryGetValue(_fontType, out Type? type) && (null != type))
            {
                var font = Element.Elements<A.TextFontType>().FirstOrDefault(e => type.IsInstanceOfType(e));
                font?.Remove();
            }
        }

        /// <summary>
        /// Get the font family as the normalized value
        /// </summary>
        /// <param name="_fontElement">The font element containing the font family name</param>
        /// <returns></returns>
        protected string? GetNormalizeFontFamily(A.TextFontType? _fontElement)
        {
            string? family = _fontElement?.Typeface;
            NormalizeFontFamily(ref family, Language, AlternativeLanguage, Element);
            return family;
        }

        /// <summary>
        /// Get the abbreviation of the font type
        /// </summary>
        /// <param name="_fontType"></param>
        /// <returns></returns>
        protected static string GetFontTypeAbbreviation(A.TextFontType _fontType)
        {
            return _fontType switch
            {
                A.ComplexScriptFont => "cs",
                A.EastAsianFont => "ea",
                _ => "lt"
            };
        }

        /// <summary>
        /// Normalize the font family
        /// </summary>
        /// <param name="_family">The family wanna be normalized</param>
        /// <param name="_lang">The language of the text</param>
        /// <param name="_altLang">The alternative language of the text</param>
        /// <param name="_belongElement">The element containing this font declaretion. This argument is used to locate the scheme</param>
        public static void NormalizeFontFamily(ref string? _family, string? _lang, string? _altLang, OpenXmlElement? _belongElement)
        {
            if (null != _family)
            {
                var match = Regex.Match(_family.Trim(), @"^\+(mn|mj)\-(lt|cs|ea)$");
                if (match.Success)
                {
                    if (null != _belongElement)
                    {
                        A.FontCollectionType? fonts = null;
                        if (match.Groups[1].Value == "mj")
                        {
                            FelisSlide.RetrospectToSlide(_belongElement)?.GetThemeScheme<A.FontScheme>(e =>
                            {
                                if (null != e.MajorFont)
                                {
                                    fonts = e.MajorFont;
                                    return true;
                                }
                                return false;
                            });
                        }
                        else
                        {
                            FelisSlide.RetrospectToSlide(_belongElement)?.GetThemeScheme<A.FontScheme>(e =>
                            {
                                if (null != e.MajorFont)
                                {
                                    fonts = e.MajorFont;
                                    return true;
                                }
                                return false;
                            });
                        }
                        _family = GetFontFamilyFromCollection(fonts, _lang, _altLang, match.Groups[2].Value);
                    }
                }
            }
        }

        /// <summary>
        /// Get the font family of the special font's type from a collection
        /// </summary>
        /// <param name="_fonts">The collection containing the fonts</param>
        /// <param name="_lang">The language of the text</param>
        /// <param name="_altLang">The alternative language of the text</param>
        /// <param name="_fontType">The special font's type</param>
        /// <returns></returns>
        private static string? GetFontFamilyFromCollection(A.FontCollectionType? _fonts, string? _lang, string? _altLang, string _fontType)
        {
            if (null != _fonts)
            {
                TextFontType? fontType = (_fontType == "cs") ? _fonts.ComplexScriptFont : ((_fontType == "ea") ? _fonts.EastAsianFont : _fonts.LatinFont);
                string? family = fontType?.Typeface;
                if (!string.IsNullOrWhiteSpace(family))
                {
                    return family;
                }
                else if (fontType is A.ComplexScriptFont)
                {
                    family = _fonts.LatinFont?.Typeface;
                    return string.IsNullOrWhiteSpace(family) ? null : family;
                }
                else if (fontType is A.EastAsianFont)
                {
                    family = null;
                    foreach (var fontItem in _fonts.Elements<A.SupplementalFont>())
                    {
                        if (fontItem.Script == _lang)
                        {
                            family = fontItem.Typeface;
                            break;
                        }
                        else if (fontItem.Script == _altLang)
                        {
                            family = fontItem.Typeface;
                        }
                    }
                    if (string.IsNullOrWhiteSpace(family))
                    {
                        family = _fonts.LatinFont?.Typeface;
                    }
                    return string.IsNullOrWhiteSpace(family) ? null : family;
                }
            }
            return null;
        }

        /// <summary>
        /// The color of the foreground
        /// </summary>
        public FelisColor? Color => FelisColor.Create(Element.GetFirstChild<A.SolidFill>() ?? new A.SolidFill(), obj =>
        {
            if ((obj is FelisColor color) && (null != color.ContainerElement) && (null == color.ContainerElement.Parent))
            {
                Element.AddChild(color.ContainerElement, false);
            }
        });
    }   
}
