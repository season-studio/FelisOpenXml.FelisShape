using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// Attribute for discribe the sub class of the FelisGraphicFrame
    /// </summary>
    internal class FelisGraphicFrameAttribute : Attribute
    {
        public readonly string DataNamespace;

        public FelisGraphicFrameAttribute(string _dataNamespace) 
        { 
            DataNamespace = _dataNamespace;
        }
    }

    /// <summary>
    /// The base class of the graphic frame shape
    /// </summary>
    [FelisShapeClass(typeof(P.GraphicFrame), 
        NonVisualDrawingPropertiesChain = new[] { typeof(P.NonVisualGraphicFrameProperties) })
    ]
    public abstract class FelisGraphicFrame : FelisShape
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_element">The graphic frame element</param>
        protected FelisGraphicFrame(P.GraphicFrame _element)
            : base(_element) 
        { 
        }

        /// <summary>
        /// Get the graphic data.
        /// An empty one will be created if there is no existed one.
        /// </summary>
        protected A.GraphicData ForceGraphicData
        {
            get
            {
                if (Element is P.GraphicFrame graphicFrame)
                {
                    if (null == graphicFrame.Graphic)
                    {
                        graphicFrame.Graphic = new A.Graphic(new A.GraphicData());
                    }
                    if (null == graphicFrame.Graphic.GraphicData)
                    {
                        graphicFrame.Graphic.GraphicData = new A.GraphicData();
                    }
                    return graphicFrame.Graphic.GraphicData;
                }
                return new A.GraphicData();
            }
        }

        internal delegate FelisGraphicFrame? CreateHandler(P.GraphicFrame _element);

        internal static readonly IReadOnlyDictionary<string, CreateHandler> CreateorMap = (new Func<IReadOnlyDictionary<string, CreateHandler>>(() =>
        {
            var map = new Dictionary<string, CreateHandler>();
            
            var ctorArgTypes = new[] { typeof(P.GraphicFrame) };

            var assembly = typeof(FelisGraphicFrame).Assembly;
            foreach (var type in assembly.GetTypes())
            {
                if (type.IsSubclassOf(typeof(FelisGraphicFrame)))
                {
                    foreach (var attr in type.GetCustomAttributes(false))
                    {
                        if (attr is FelisGraphicFrameAttribute eAttr)
                        {
                            var ctorInfo = type.GetConstructor(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.CreateInstance | BindingFlags.Instance, ctorArgTypes);
                            if (null != ctorInfo)
                            {
                                CreateHandler creator = (P.GraphicFrame _element) => (ctorInfo.Invoke(new[] { _element }) as FelisGraphicFrame);
                                if (null != creator)
                                {
                                    map[eAttr.DataNamespace] = creator;
                                }
                            }
                        }
                    }
                }
            }

            return map;
        }))();

        internal static FelisGraphicFrame? FromElement(P.GraphicFrame _element)
        {
            string? dataUri = _element.Graphic?.GraphicData?.Uri;
            if ((null != dataUri) && CreateorMap.TryGetValue(dataUri, out CreateHandler? _creator))
            {
                return _creator(_element);
            }

            return new FelisUnknownGraphicFrame(_element);
        }
    }

    /// <summary>
    /// The class for unknown graphic frame shape
    /// </summary>
    public class FelisUnknownGraphicFrame : FelisGraphicFrame
    {
        internal FelisUnknownGraphicFrame(P.GraphicFrame _element)
            : base(_element)
        {
            
        }
    }
}
