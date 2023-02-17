using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing;
using FelisOpenXml.FelisShape.Base;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// The basic class of the blip element
    /// </summary>
    /// <typeparam name="T">The type of the blip element</typeparam>
    public abstract class FelisBlipBase<T> : FelisCompositeElement
        where T : OpenXmlCompositeElement
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_element">The blip element</param>
        protected FelisBlipBase(T _element)
            : base(_element) 
        { 
        }

        private A.FillRectangle ForceStretchRect
        {
            get
            {
                var stretch = Element.GetFirstChild<A.Stretch>();
                if (null == stretch)
                {
                    stretch = Element.AppendChild(new A.Stretch());
                }
                var rect = stretch.GetFirstChild<A.FillRectangle>();
                if (null == rect)
                {
                    rect = stretch.AppendChild(new A.FillRectangle());
                }
                return rect;
            }
        }

        private A.SourceRectangle ForceSourceRect
        {
            get
            {
                return Element.GetFirstChild<A.SourceRectangle>() ?? Element.AppendChild(new A.SourceRectangle());
            }
        }

        /// <summary>
        /// Get the stretch object
        /// </summary>
        public FelisRelativeRect<A.FillRectangle> Stretch => new FelisRelativeRect<A.FillRectangle>(ForceStretchRect);

        /// <summary>
        /// Get the displacement of the source
        /// </summary>
        public FelisRelativeRect<A.SourceRectangle> SourceDisplacement => new FelisRelativeRect<SourceRectangle>(ForceSourceRect);

        /// <summary>
        /// Get the data of the image
        /// </summary>
        /// <param name="_buffer">The buffer to receive the data. This argument can be a stream or an array of byte.</param>
        public void CopyTo(object _buffer)
        {
            var blip = Element.GetFirstChild<A.Blip>();
            string? resId = blip?.Embed;
            if (!string.IsNullOrWhiteSpace(resId))
            {
                var slidePart = FelisSlide.RetrospectToSlideElement(blip)?.SlidePart;
                var stream = slidePart?.GetPartById(resId)?.GetStream();
                if (null != stream)
                {
                    using (stream)
                    {
                        if ((_buffer is Stream targetStream) && (targetStream.CanWrite))
                        {
                            stream.CopyTo(targetStream);
                        }
                        else if (_buffer is byte[] byteBuf)
                        {
                            stream.Read(byteBuf);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Set a new image to this blip
        /// </summary>
        /// <param name="_source">The buffer containing the image. This argument can be a stream or an array of byte</param>
        /// <param name="_type">The type of the image. Such as "Bmp", "Png", and so on.</param>
        public void Set(object _source, string? _type)
        {
            var blip = Element.GetFirstChild<A.Blip>();
            if (null == blip)
            {
                blip = Element.InsertAt(new A.Blip(), 0);
            }

            string? resId = blip.Embed;
            var slidePart = FelisSlide.RetrospectToSlideElement(blip)?.SlidePart;
            if (null != slidePart)
            {
                if (null != resId)
                {
                    slidePart.DeletePart(resId);
                    blip.Embed = string.Empty;
                }
                
                ImagePartType imagePartType;
                if (!Enum.TryParse(_type, true, out imagePartType))
                {
                    imagePartType = default;
                }
                var imagePart = slidePart?.AddImagePart(imagePartType);
                if (null != imagePart)
                {
                    resId = slidePart?.GetIdOfPart(imagePart);
                    blip.Embed = resId;

                    if (_source is Stream sourceStream)
                    {
                        imagePart.FeedData(sourceStream);
                    }
                    else if (_source is byte[] sourceBytes)
                    {
                        using (var stream = new MemoryStream(sourceBytes))
                        {
                            imagePart.FeedData(stream);
                        }
                    }
                }

                var dropExts = blip.GetFirstChild<A.BlipExtensionList>()?.ChildElements.Where(e => e.ChildElements.Where(e2 => e2.LocalName.EndsWith("blip", StringComparison.OrdinalIgnoreCase)).Any()).ToArray();
                if (null != dropExts)
                {
                    foreach (var dropItem in dropExts)
                    {
                        dropItem.Remove();
                    }
                }
            }
        }
    }

    /// <summary>
    /// The class for operating the relative displacement
    /// </summary>
    public class FelisRelativeRect<T>
        where T : A.RelativeRectangleType
    {
        /// <summary>
        /// The element containing the rectangle's information
        /// </summary>
        public readonly T RectElement;

        internal FelisRelativeRect(T _element)
        {
            RectElement = _element;
        }

        /// <summary>
        /// Left
        /// </summary>
        public double? Left
        {
            get
            {
                return RectElement.Left?.Value / 100000 ?? 0;
            }

            set
            {
                if (value is null)
                {
                    RectElement.Left = null;
                }
                else
                {
                    RectElement.Left = (int)(value * 100000);
                }
            }
        }

        /// <summary>
        /// Right
        /// </summary>
        public double? Right
        {
            get
            {
                return RectElement.Right?.Value / 100000 ?? 0;
            }

            set
            {
                if (value is null)
                {
                    RectElement.Right = null;
                }
                else
                {
                    RectElement.Right = (int)(value * 100000);
                }
            }
        }

        /// <summary>
        /// Top
        /// </summary>
        public double? Top
        {
            get
            {
                return RectElement.Top?.Value / 100000 ?? 0;
            }

            set
            {
                if (value is null)
                {
                    RectElement.Top = null;
                }
                else
                {
                    RectElement.Top = (int)(value * 100000);
                }
            }
        }

        /// <summary>
        /// Bottom
        /// </summary>
        public double? Bottom
        {
            get
            {
                return RectElement.Bottom?.Value / 100000 ?? 0;
            }

            set
            {
                if (value is null)
                {
                    RectElement.Bottom = null;
                }
                else
                {
                    RectElement.Bottom = (int)(value * 100000);
                }
            }
        }

        /// <summary>
        /// Clear all the setting of the displacement
        /// </summary>
        public void Clear()
        {
            RectElement.ClearAllAttributes();
        }

        /// <summary>
        /// Remove the element from it's parent.
        /// This is a danger operating. After invoking this method, all the changing of the displacement have no effect in the shape.
        /// </summary>
        public void Delete()
        {
            RectElement.Remove();
        }
    }
}
