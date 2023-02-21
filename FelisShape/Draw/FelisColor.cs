using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using FelisOpenXml.FelisShape.Base;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// The color class for the shape, text, etc.
    /// </summary>
    public class FelisColor : FelisUnderlingElement
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_container">The element containing the color element</param>
        /// <param name="_submitter">The action invoked after changing.</param>
        protected FelisColor(OpenXmlCompositeElement _container, Action<object>? _submitter)
            : base(_container, _submitter)
        {

        }

        /// <summary>
        /// Reload the color element
        /// </summary>
        protected override void Reload()
        {
            workElement = GetColorElement(Element);
        }

        /// <summary>
        /// The RGB value of the color
        /// </summary>
        public Color Value
        {
            get
            {
                return ConvertToColor(workElement);
            }

            set
            {
                foreach (var item in Element.ChildElements.Where(e => (e is A.HslColor) || (e is A.SchemeColor) || (e is A.SystemColor) || (e is A.PresetColor)).ToArray())
                {
                    item.Remove();
                }
                if (workElement is A.RgbColorModelHex rgbHex)
                {
                    rgbHex.Val = $"{value.R.ToString("X2")}{value.G.ToString("X2")}{value.B.ToString("X2")}";
                }
                else if (workElement is A.RgbColorModelPercentage rgbPortion)
                {
                    rgbPortion.RedPortion = value.R;
                    rgbPortion.GreenPortion = value.G;
                    rgbPortion.BluePortion = value.B;
                }
                else
                {
                    Element.AddChild(new A.RgbColorModelHex()
                    {
                        Val = $"{value.R.ToString("X2")}{value.G.ToString("X2")}{value.B.ToString("X2")}"
                    }, false);
                }
                Submit();
            }
        }

        /// <summary>
        /// Set the color by HSL format
        /// </summary>
        /// <param name="_h"></param>
        /// <param name="_s"></param>
        /// <param name="_l"></param>
        public void SetAsHSL(int _h, int _s, int _l)
        {
            foreach (var item in Element.ChildElements.Where(e => (e is A.RgbColorModelHex) || (e is A.RgbColorModelPercentage) || (e is A.SchemeColor) || (e is A.SystemColor) || (e is A.PresetColor)).ToArray())
            {
                item.Remove();
            }

            if (workElement is A.HslColor hsl)
            {
                hsl.HueValue = _h;
                hsl.SatValue = _s;
                hsl.LumValue = _l;
            }
            else
            {
                Element.AddChild(new A.HslColor()
                {
                    HueValue = _h,
                    SatValue = _s,
                    LumValue = _l
                }, false);
            }
            Submit();
        }

        /// <summary>
        /// Set the color to use the scheme
        /// </summary>
        /// <param name="_color"></param>
        public void SetAsScheme(A.SchemeColorValues _color)
        {
            foreach (var item in Element.ChildElements.Where(e => (e is A.RgbColorModelHex) || (e is A.RgbColorModelPercentage) || (e is A.HslColor) || (e is A.SystemColor) || (e is A.PresetColor)).ToArray())
            {
                item.Remove();
            }

            if (workElement is A.SchemeColor schemeColor)
            {
                schemeColor.Val = _color;
            }
            else
            {
                Element.AddChild(new A.SchemeColor()
                {
                    Val = _color
                }, false);
            }
            Submit();
        }

        /// <summary>
        /// Clear the color
        /// </summary>
        public void Clear()
        {
            foreach (var item in Element.ChildElements.Where(e => (e is A.RgbColorModelHex) || (e is A.RgbColorModelPercentage) || (e is A.SchemeColor) || (e is A.HslColor) || (e is A.SystemColor) || (e is A.PresetColor)).ToArray())
            {
                item.Remove();
            }
            Submit();
        }

        /// <summary>
        /// Check if the color value is defined.
        /// </summary>
        public bool HasDefined => Element.ChildElements.Where(e => (e is A.RgbColorModelHex) || (e is A.RgbColorModelPercentage) || (e is A.SchemeColor) || (e is A.HslColor) || (e is A.SystemColor) || (e is A.PresetColor)).Any();

        /// <summary>
        /// Create an instance of the color object
        /// </summary>
        /// <param name="_container">The element which contains the color element</param>
        /// <param name="_fnSubmit">Action on submitting the changings. It is often used to add the new container into the DOM tree.</param>
        /// <returns></returns>
        internal static FelisColor? Create(OpenXmlCompositeElement? _container, Action<object>? _fnSubmit = null)
        {
            return null == _container ? null : FSUtilities.SingletonAssignObject<FelisColor>(_container, (_) =>
            {
                return new FelisColor(_container, _fnSubmit);
            });
        }

        /// <summary>
        /// Get the color element in the special container
        /// </summary>
        /// <param name="_element">The element of the container</param>
        /// <returns></returns>
        public static OpenXmlElement? GetColorElement(OpenXmlElement? _element)
        {
            return _element?.ChildElements.LastOrDefault(e =>
            {
                return (e is A.RgbColorModelHex)
                        || (e is A.RgbColorModelPercentage)
                        || (e is A.HslColor)
                        || (e is A.SchemeColor)
                        || (e is A.SystemColor)
                        || (e is A.PresetColor);
            });
        }

        /// <summary>
        /// The fallback value of the color.
        /// It is used when some exception raised.
        /// </summary>
        internal static readonly Color FallbackColor = Color.Transparent;

        /// <summary>
        /// Convert the value in a special color element to the type of Color
        /// </summary>
        /// <param name="_colorElement">The special color element</param>
        /// <param name="_disableScheme">True for forbidden searching the color scheme.</param>
        /// <returns></returns>
        public static Color ConvertToColor(OpenXmlElement? _colorElement, bool _disableScheme = false)
        {
            return AdjustColor(_colorElement, _colorElement switch
            {
                A.RgbColorModelHex rgbHex => ConvertToColorFromRGBHex(rgbHex),
                A.RgbColorModelPercentage rgbPortion => ConvertToColorFromRGBPortion(rgbPortion),
                A.HslColor hslPortion => ConvertToColorFromHSL(hslPortion),
                A.SchemeColor schemeColor => _disableScheme ? FallbackColor : ConvertToColorFromScheme(schemeColor),
                A.SystemColor => FallbackColor,
                A.PresetColor => FallbackColor,
                _ => FallbackColor
            });
        }

        /// <summary>
        /// Adjust the color according to the information conatined in the child element of the color element
        /// </summary>
        /// <param name="_colorElement">The color element</param>
        /// <param name="_color">The origin color</param>
        /// <returns></returns>
        public static Color AdjustColor(OpenXmlElement? _colorElement, Color _color)
        {
            // TODO:
            return _color;
        }

        /// <summary>
        /// Get the color map of the given slide
        /// </summary>
        /// <param name="_slidePart">The part of the special slide</param>
        /// <returns></returns>
        public static OpenXmlCompositeElement? GetColorMap(SlidePart? _slidePart)
        {
            if (null == _slidePart)
            {
                return null;
            }

            var masterColorMap = _slidePart.SlideLayoutPart?.SlideMasterPart?.SlideMaster.ColorMap;

            //从当前Slide获取ColorMap
            if (_slidePart.Slide.ColorMapOverride != null)
            {
                if (_slidePart.Slide.ColorMapOverride.MasterColorMapping != null)
                {
                    return masterColorMap;
                }

                if (_slidePart.Slide.ColorMapOverride.OverrideColorMapping != null)
                {
                    return _slidePart.Slide.ColorMapOverride.OverrideColorMapping;
                }
            }

            //从SlideLayout获取ColorMap
            if (_slidePart.SlideLayoutPart?.SlideLayout.ColorMapOverride != null)
            {
                if (_slidePart.SlideLayoutPart.SlideLayout.ColorMapOverride.MasterColorMapping != null)
                {
                    return masterColorMap;
                }

                if (_slidePart.SlideLayoutPart.SlideLayout.ColorMapOverride.OverrideColorMapping != null)
                {
                    return _slidePart.SlideLayoutPart.SlideLayout.ColorMapOverride.OverrideColorMapping;
                }
            }

            //从SlideMaster获取ColorMap
            return masterColorMap;
        }

        /// <summary>
        /// Convert a scheme color to the type of Color
        /// </summary>
        /// <param name="_element">The element of the color as the scheme format</param>
        /// <returns></returns>
        public static Color ConvertToColorFromScheme(A.SchemeColor _element)
        {
            var slide = FelisSlide.RetrospectToSlide(_element);
            if (null != slide)
            {
                var colorMap = GetColorMap(slide.SDKPart);
                var colorScheme = slide.GetThemeScheme<A.ColorScheme>();
                var colorKey = ((EnumValue<A.SchemeColorValues>)(_element.Val?.Value ?? A.SchemeColorValues.Accent1)).ToString();
                if (null != colorMap)
                {
                    colorKey = (colorMap.GetAttributes().FirstOrDefault(e => e.LocalName == colorKey).Value ?? colorKey);
                }

                if (!string.IsNullOrEmpty(colorKey) && (null != colorScheme))
                {
                    var colorType = colorScheme.ChildElements.FirstOrDefault(e => (e is A.Color2Type) && (e.LocalName == colorKey)) as A.Color2Type;
                    return ConvertToColor(GetColorElement(colorType), true);
                }
            }

            return FallbackColor;
        }

        /// <summary>
        /// Convert the HSL color to the type of Color
        /// </summary>
        /// <param name="_element">The element of the color as the HSL format</param>
        /// <returns></returns>
        public static Color ConvertToColorFromHSL(A.HslColor _element)
        {
            try
            {
                int oriH = (_element.HueValue?.Value ?? 0);
                double nmH = oriH * 359 / 255;
                int oriS = (_element.SatValue?.Value ?? 0);
                double nmS = oriS / 255;
                int oriL = (_element.LumValue?.Value ?? 0);
                int nmL = oriL / 255;


                if (0 == oriS)
                {
                    var v = (byte)(oriL);
                    return Color.FromArgb(v, v, v);
                }
                else
                {
                    double C = (1 - Math.Abs(2 * nmL - 1)) * nmS;
                    double hh = nmH / 60.0;
                    double X = C * (1 - Math.Abs(hh % 2 - 1));
                    double r = 0, g = 0, b = 0;
                    if (hh >= 0 && hh < 1)
                    {
                        r = C;
                        g = X;
                    }
                    else if (hh >= 1 && hh < 2)
                    {
                        r = X;
                        g = C;
                    }
                    else if (hh >= 2 && hh < 3)
                    {
                        g = C;
                        b = X;
                    }
                    else if (hh >= 3 && hh < 4)
                    {
                        g = X;
                        b = C;
                    }
                    else if (hh >= 4 && hh < 5)
                    {
                        r = X;
                        b = C;
                    }
                    else
                    {
                        r = C;
                        b = X;
                    }
                    double m = nmL - C / 2;
                    r += m;
                    g += m;
                    b += m;
                    r *= 255.0;
                    g *= 255.0;
                    b *= 255.0;
                    r = Math.Round(r);
                    g = Math.Round(g);
                    b = Math.Round(b);
                    return Color.FromArgb((int)r, (int)g, (int)b);
                }
            }
            catch (Exception err)
            {
                Trace.TraceWarning(err.ToString());
            }

            return FallbackColor;
        }

        /// <summary>
        /// Convert the RGB portion values to color
        /// </summary>
        /// <param name="_element">The element containing the RGB portions</param>
        /// <returns></returns>
        public static Color ConvertToColorFromRGBPortion(A.RgbColorModelPercentage _element)
        {
            try
            {
                return Color.FromArgb(_element.RedPortion?.Value ?? 0, _element.GreenPortion?.Value ?? 0, _element.BluePortion?.Value ?? 0);
            }
            catch (Exception err)
            {
                Trace.TraceWarning(err.ToString());
            }

            return FallbackColor;
        }

        /// <summary>
        /// Convert the RGB value written by hex string o the color 
        /// </summary>
        /// <param name="_element">The element containing the RGB value written by hex string</param>
        /// <returns></returns>
        public static Color ConvertToColorFromRGBHex(A.RgbColorModelHex _element)
        {
            try
            {
                var hexStr = _element.Val?.Value?.Trim();
                if (!string.IsNullOrEmpty(hexStr))
                {
                    int end = hexStr.Length;
                    int start = Math.Max(end - 2, 0);
                    var bHex = hexStr.Substring(start, end - start);
                    var bVal = byte.Parse(bHex, System.Globalization.NumberStyles.HexNumber);
                    end = start;
                    start = Math.Max(end - 2, 0);
                    var gHex = hexStr.Substring(start, end - start);
                    var gVal = byte.Parse(gHex, System.Globalization.NumberStyles.HexNumber);
                    end = start;
                    start = Math.Max(end - 2, 0);
                    var rHex = hexStr.Substring(start, end - start);
                    var rVal = byte.Parse(rHex, System.Globalization.NumberStyles.HexNumber);
                    return Color.FromArgb(rVal, gVal, bVal);
                }
            }
            catch (Exception err)
            {
                Trace.TraceWarning(err.ToString());
            }

            return FallbackColor;
        }
    }
}
