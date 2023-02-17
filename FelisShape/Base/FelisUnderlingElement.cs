using DocumentFormat.OpenXml;
using FelisOpenXml.FelisShape.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The base class for the element contained in another.
    /// This class is used for the element which's container may not be connected to the DOM, or for the element which have no specially type.
    /// </summary>
    public abstract class FelisUnderlingElement : FelisCompositeElement
    {
        /// <summary>
        /// The element this object manipulates
        /// </summary>
        protected OpenXmlElement? workElement;
        /// <summary>
        /// The action invoked after changing the element
        /// </summary>
        protected readonly Action<object>? Submitter;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_container">The element containing the working element for this object maipulating.</param>
        /// <param name="_submitter">The action invoked after changing the element</param>
        protected FelisUnderlingElement(OpenXmlCompositeElement _container, Action<object>? _submitter = null)
            : base(_container)
        {
            Submitter = _submitter;
            Reload();
        }

        /// <summary>
        /// The element of the container. It is as the same as the Element property.
        /// </summary>
        public OpenXmlCompositeElement ContainerElement => Element;

        /// <summary>
        /// The element of the working element for the target information of this object.
        /// </summary>
        public OpenXmlElement? WorkElement => workElement;

        /// <summary>
        /// Reload the working element
        /// </summary>
        protected abstract void Reload();

        /// <summary>
        /// Submit the changing of the information contained in this object
        /// </summary>
        protected virtual void Submit()
        {
            Submitter?.Invoke(this);
            Reload();
        }
    }
}
