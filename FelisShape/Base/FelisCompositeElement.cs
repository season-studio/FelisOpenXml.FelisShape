using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using System.Reflection;

namespace FelisOpenXml.FelisShape.Base
{
    /// <summary>
    /// The basic class for the composite element
    /// </summary>
    public abstract class FelisCompositeElement : IEquatable<FelisCompositeElement>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_element">The composite element</param>
        /// <exception cref="ArgumentNullException">The input argument is null</exception>
        protected FelisCompositeElement(OpenXmlCompositeElement? _element)
        {
            if (null == _element)
            {
                throw new ArgumentNullException(nameof(_element));
            }
            Element = _element;
        }

        /// <summary>
        /// The Open Xml Element assigned to this object. Using this member for invoking the API of the OOXML.
        /// </summary>
        public OpenXmlCompositeElement Element { get; private set; }

        #region Function for equal operating
        /// <summary>
        /// Test the object is equal to another 
        /// </summary>
        /// <param name="other">The other object</param>
        /// <returns></returns>
        public bool Equals(FelisCompositeElement? other)
        {
            return (other?.Element == Element);
        }

        /// <summary>
        /// Test the object is equal to another 
        /// </summary>
        /// <param name="_value1">The first object</param>
        /// <param name="_value2">The second object</param>
        /// <returns></returns>
        public static bool operator ==(FelisCompositeElement? _value1, FelisCompositeElement? _value2)
        {
            return object.Equals(_value1, _value2);
        }

        /// <summary>
        /// Test the object is not equal to another 
        /// </summary>
        /// <param name="_value1">The first object</param>
        /// <param name="_value2">The second object</param>
        /// <returns></returns>
        public static bool operator !=(FelisCompositeElement? _value1, FelisCompositeElement? _value2)
        {
            return !object.Equals(_value1, _value2);
        }

        /// <summary>
        /// Test the object is equal to another 
        /// </summary>
        /// <param name="_obj">The other object</param>
        /// <returns></returns>
        public override bool Equals(object? _obj)
        {
            return (_obj is FelisCompositeElement other) && EqualityComparer<FelisCompositeElement?>.Default.Equals(this, other);
        }

        /// <summary>
        /// Get the hash code of the object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
        #endregion
    }
}