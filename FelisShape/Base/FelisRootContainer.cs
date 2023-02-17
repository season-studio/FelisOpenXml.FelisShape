using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FelisOpenXml.FelisShape.Base
{
    /// <summary>
    /// The basic class for the root element in a part
    /// </summary>
    /// <typeparam name="TPart"></typeparam>
    /// <typeparam name="TRoot"></typeparam>
    public abstract class FelisRootContainer<TPart, TRoot> : IEquatable<FelisRootContainer<TPart, TRoot>>
        where TPart : OpenXmlPart
        where TRoot : OpenXmlCompositeElement
    {
        /// <summary>
        /// The part containing the root element
        /// </summary>
        public readonly TPart SDKPart;
        /// <summary>
        /// The root element
        /// </summary>
        public readonly TRoot SDKRootElement;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_part">The part containing the root element</param>
        /// <param name="_element">The root element</param>
        /// <exception cref="ArgumentNullException">Any one of the argument is null.</exception>
        protected FelisRootContainer(TPart? _part, TRoot? _element) 
        {
            if (null == _part)
            {
                throw new ArgumentNullException(nameof(_part));
            }
            if (null == _element) 
            {
                throw new ArgumentNullException(nameof(_element));
            }
            SDKPart = _part;
            SDKRootElement = _element;
        }

        /// <summary>
        /// Test the object is equal to another 
        /// </summary>
        /// <param name="other">The other object</param>
        /// <returns></returns>
        public bool Equals(FelisRootContainer<TPart, TRoot>? other)
        {
            return (other?.SDKPart == SDKPart);
        }

        /// <summary>
        /// Test the object is equal to another 
        /// </summary>
        /// <param name="_value1">The first object</param>
        /// <param name="_value2">The second object</param>
        /// <returns></returns>
        public static bool operator ==(FelisRootContainer<TPart, TRoot>? _value1, FelisRootContainer<TPart, TRoot>? _value2)
        {
            return object.Equals(_value1, _value2);
        }

        /// <summary>
        /// Test the object is not equal to another 
        /// </summary>
        /// <param name="_value1">The first object</param>
        /// <param name="_value2">The second object</param>
        /// <returns></returns>
        public static bool operator !=(FelisRootContainer<TPart, TRoot>? _value1, FelisRootContainer<TPart, TRoot>? _value2)
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
            return (_obj is FelisRootContainer<TPart, TRoot> other) && EqualityComparer<FelisRootContainer<TPart, TRoot>?>.Default.Equals(this, other);
        }

        /// <summary>
        /// Get the hash code of the object
        /// </summary>
        /// <returns></returns>
        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}
