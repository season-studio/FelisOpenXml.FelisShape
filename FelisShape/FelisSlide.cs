using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using System.Xml.Linq;
using System.Xml;
using System.Diagnostics;
using FelisOpenXml.FelisShape.Base;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The class of the slide
    /// </summary>
    public class FelisSlide : FelisRootContainer<SlidePart, P.Slide>, IFelisShapeTree
    {
        /// <summary>
        /// The presentation which contains this slide
        /// </summary>
        public readonly FelisPresentation Presentation;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_presentation">The presentation which contains this slide</param>
        /// <param name="_part">The part of the slide</param>
        protected FelisSlide(FelisPresentation _presentation, SlidePart _part)
            : base(_part, _part.Slide)
        {
            Presentation = _presentation;
        }

        /// <summary>
        /// Create a new instance or get an existed instance of the slide
        /// </summary>
        /// <param name="_presentation">The presentation which contains the slide</param>
        /// <param name="_part">The part of the slide</param>
        /// <returns></returns>
        public static FelisSlide? From(FelisPresentation? _presentation, SlidePart _part)
        {
            return FSUtilities.SingletonAssignObject<FelisSlide>(_part, (_) =>
            {
                return (null == _presentation) ? null : new FelisSlide(_presentation, _part);
            });
        }

        /// <summary>
        /// Check if there is any shape in the slide
        /// </summary>
        public bool HasShapes => FelisShapeGroup.CheckHasShapes(SDKRootElement.CommonSlideData?.ShapeTree);

        /// <summary>
        /// Get an iterator of the shapes in the slide
        /// </summary>
        public IEnumerable<FelisShape> Shapes => FelisShapeGroup.GetChildrenShapes(SDKRootElement.CommonSlideData?.ShapeTree);

        /// <summary>
        /// Submit all the changing in the slide
        /// </summary>
        public void Submit()
        {
            SDKRootElement.Save();
        }

        /// <summary>
        /// Remove the slide
        /// </summary>
        public void Remove()
        {
            Presentation.RemoveSlide(this);
        }

        /// <summary>
        /// Get an iterator of the customer data parts assigned to the slide
        /// </summary>
        public IEnumerable<OpenXmlPart> CustomerDataParts
        {
            get
            {
                var customerDataList = SDKRootElement.CommonSlideData?.CustomerDataList;
                var customDatas = customerDataList?.Elements<P.CustomerData>();
                if (null != customDatas)
                {
                    foreach (var customData in customDatas)
                    {
                        var customPart = SDKPart.GetPartById(customData?.Id?.Value ?? string.Empty);
                        if (null != customPart)
                        {
                            yield return customPart;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Add a new customer data part assigned to the slide
        /// </summary>
        /// <param name="_fnInit">The function for initializing the customer data</param>
        /// <returns>The instance of the new customer data part.</returns>
        public OpenXmlPart? AddCustomerData(Func<OpenXmlPart, bool>? _fnInit = null)
        {
            var part = Presentation.SDKPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            if (null != part)
            {
                bool passInit;
                try
                {
                    passInit = _fnInit?.Invoke(part) ?? true;

                    if (passInit)
                    {
                        var rid = SDKPart.CreateRelationshipToPart(part);
                        if (!string.IsNullOrWhiteSpace(rid))
                        {
                            var cData = SDKRootElement.CommonSlideData;
                            if (null == cData)
                            {
                                cData = (SDKRootElement.CommonSlideData = new P.CommonSlideData());
                            }
                            var cList = cData.CustomerDataList;
                            if (null == cList)
                            {
                                cList = (cData.CustomerDataList = new P.CustomerDataList());
                            }
                            var custom = new P.CustomerData()
                            {
                                Id = rid
                            };
                            var refCustomItem = cList.GetFirstChild<P.CustomerDataTags>();
                            if (null != refCustomItem)
                            {
                                cList.InsertBefore(custom, refCustomItem);
                            }
                            else
                            {
                                cList.InsertAt(custom, 0);
                            }
                        }
                    }
                }
                catch (Exception err)
                {
                    Trace.TraceWarning(err.ToString());
                    passInit = false;
                }

                if (!passInit)
                {
                    if (null != part)
                    {
                        Presentation.SDKPart.DeletePart(part);
                        part = null;
                    }
                }

                SDKRootElement.Save();
            }

            return part;
        }

        /// <summary>
        /// Delete a special custom data part
        /// </summary>
        /// <param name="_part"></param>
        /// <returns></returns>
        public int RemoveCustomDataPart(OpenXmlPart? _part)
        {
            if (_part is CustomXmlPart)
            {
                try
                {
                    var rId = SDKPart.GetIdOfPart(_part);
                    var cList = SDKRootElement.CommonSlideData?.CustomerDataList;
                    if (null != cList)
                    {
                        foreach (var item in cList.Elements<P.CustomerData>())
                        {
                            if (item.Id == rId)
                            {
                                item.Remove();
                                break;
                            }
                        }
                    }
                    return SDKPart.DeletePart(_part) ? 0 : -3;
                }
                catch (ArgumentOutOfRangeException)
                {
                    return -2;
                }
            }
            return -1;
        }

        internal static readonly string CustomerDataNameSpace = "https://www.seson-studio.top/felis-ooxml/custom-data";
        internal static readonly string CustomerDataPrefix = "fcd";
        internal static readonly string CustomerDataIDAttr = $"{CustomerDataPrefix}:id";

        /// <summary>
        /// Check if there is a customer data with the special id in the slide
        /// </summary>
        /// <param name="_id">The special id of the customer data</param>
        /// <returns></returns>
        public bool CheckHasCustomerData(string _id)
        {
            foreach (var part in CustomerDataParts)
            {
                if (part is CustomXmlPart xmlPart)
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    using (var stream = xmlPart.GetStream())
                    {
                        xmlDoc.Load(stream);
                        if ((xmlDoc.DocumentElement?.NamespaceURI == CustomerDataNameSpace)
                             && (xmlDoc.DocumentElement?.Attributes[CustomerDataIDAttr]?.Value == _id))
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Remove a customer data with the special id in the slide
        /// </summary>
        /// <param name="_id">The special id of the customer data</param>
        public void RemoveCustomerData(string _id)
        {
            foreach (var part in CustomerDataParts)
            {
                if (part is CustomXmlPart xmlPart)
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    using (var stream = xmlPart.GetStream())
                    {
                        xmlDoc.Load(stream);
                        if ((xmlDoc.DocumentElement?.NamespaceURI == CustomerDataNameSpace)
                             && (xmlDoc.DocumentElement?.Attributes[CustomerDataIDAttr]?.Value == _id))
                        {
                            var rid = SDKPart.GetIdOfPart(part);
                            var customerDataList = SDKRootElement.CommonSlideData?.CustomerDataList;
                            var customDatas = customerDataList?.Elements<P.CustomerData>();
                            if (null != customDatas)
                            {
                                foreach (var customData in customDatas)
                                {
                                    if (customData.Id == rid)
                                    {
                                        customData.Remove();
                                        break;
                                    }
                                }
                            }
                            SDKPart.DeletePart(part);
                            return;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Work with the customer data with the special id in the slide
        /// </summary>
        /// <param name="_id">The special id of the customer data</param>
        /// <param name="_fn">The action for working with the customer data. This action return true for saving the changings in the customer data.</param>
        /// <param name="_forceOne">Set true for create a temp data if the special data is not existed.</param>
        public void WorkWithCustomerData(string _id, Func<XmlDocument, bool> _fn, bool _forceOne = false)
        {
            foreach (var part in CustomerDataParts)
            {
                if (part is CustomXmlPart xmlPart)
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    using (var stream = xmlPart.GetStream())
                    {
                        xmlDoc.Load(stream);
                        if ((xmlDoc.DocumentElement?.NamespaceURI == CustomerDataNameSpace)
                             && (xmlDoc.DocumentElement?.Attributes[CustomerDataIDAttr]?.Value == _id))
                        {
                            if (_fn(xmlDoc) && stream.CanWrite)
                            {
                                stream.Seek(0, SeekOrigin.Begin);
                                stream.SetLength(0);
                                xmlDoc.Save(stream);
                            }
                            return;
                        }
                    }
                }
            }

            if (_forceOne)
            {
                XmlDocument newDoc = new XmlDocument();
                using (var newStream = new MemoryStream())
                {
                    using (var writer = XmlWriter.Create(newStream))
                    {
                        writer.WriteStartDocument(true);
                        writer.WriteStartElement(CustomerDataPrefix, "CustomData", CustomerDataNameSpace);
                        writer.WriteAttributeString(CustomerDataPrefix, "id", CustomerDataNameSpace, _id);
                        writer.WriteEndElement();
                        writer.WriteEndDocument();
                    }

                    newStream.Seek(0, SeekOrigin.Begin);
                    newDoc.Load(newStream);
                }
                if (_fn(newDoc))
                {
                    AddCustomerData((part) =>
                    {
                        using (var pStream = part.GetStream())
                        {
                            newDoc.Save(pStream);
                        }
                        return true;
                    });
                }
            }
        }

        /// <summary>
        /// Check if the slide is in a writable presentation
        /// </summary>
        public bool Writable => ((SDKPart.OpenXmlPackage.Package.FileOpenAccess & FileAccess.Write) != 0);

        /// <summary>
        /// Get the slide element which contains the given element.
        /// This function will not create a slide element if there is not an existed one.
        /// </summary>
        /// <param name="_element">The given element which should be contained in the target slide</param>
        /// <returns></returns>
        public static P.Slide? RetrospectToSlideElement(OpenXmlElement? _element)
        {
            if (null != _element)
            {
                OpenXmlElement root = _element;
                for (OpenXmlElement? scan = root; null != scan; scan = scan.Parent)
                {
                    root = scan;
                }
                return root as P.Slide;
            }
            return null;
        }

        /// <summary>
        /// Get the slide object which contains the given element.
        /// This function will not create a slide if there is not an existed one.
        /// </summary>
        /// <param name="_element">The given element which should be contained in the target slide</param>
        /// <returns></returns>
        public static FelisSlide? RetrospectToSlide(OpenXmlElement? _element)
        {
            var slidePart = RetrospectToSlideElement(_element)?.SlidePart;
            return (null == slidePart) ? null : FSUtilities.SingletonAssignObject<FelisSlide>(slidePart, null);
        }

        /// <summary>
        /// Get a special scheme in this slide
        /// </summary>
        /// <typeparam name="T">The tyep of the special scheme</typeparam>
        /// <param name="_skipSlide">Set true to scan the layout as the beginning instead of the slide</param>
        /// <returns></returns>
        public T? GetThemeScheme<T>(bool _skipSlide = false)
            where T : OpenXmlElement
        {
            var target = _skipSlide ? null : SDKPart.ThemeOverridePart?.ThemeOverride?.GetFirstChild<T>();
            if (null == target)
            {
                var layoutPart = SDKPart.SlideLayoutPart;
                if (null != layoutPart)
                {
                    target = layoutPart.ThemeOverridePart?.ThemeOverride?.GetFirstChild<T>();

                    if (null == target)
                    {
                        target = layoutPart.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.GetFirstChild<T>();
                    }
                }
            }

            return target;
        }

        /// <summary>
        /// The a special scheme assigned to this slide
        /// </summary>
        /// <typeparam name="T">The type of the target scheme</typeparam>
        /// <param name="_fnFilter">The filter for checking the scheme</param>
        /// <returns></returns>
        public T? GetThemeScheme<T>(Func<T, bool> _fnFilter)
            where T : OpenXmlElement
        {
            var target = SDKPart.ThemeOverridePart?.ThemeOverride?.GetFirstChild<T>();
            if ((null != target) && _fnFilter(target))
            {
                return target;
            }

            var layoutPart = SDKPart.SlideLayoutPart;
            if (null != layoutPart)
            {
                target = layoutPart.ThemeOverridePart?.ThemeOverride?.GetFirstChild<T>();
                if ((null != target) && _fnFilter(target))
                {
                    return target;
                }

                target = layoutPart.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.GetFirstChild<T>();
                if ((null != target) && _fnFilter(target))
                {
                    return target;
                }
            }

            return null;
        }
    }
}
