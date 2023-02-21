using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using DocumentFormat.OpenXml.Validation;

namespace FelisOpenXml.FelisShape
{   
    /// <summary>
    /// Utilities
    /// </summary>
    public static class FSUtilities
    {
        /// <summary>
        /// Copy a part
        /// </summary>
        /// <typeparam name="T">Special a type of OpenXmlPart</typeparam>
        /// <param name="_srcPart">The source part</param>
        /// <param name="_fnCreatePart">The function creating an instance of the new part</param>
        /// <returns>The new part or null</returns>
        public static T? CopyPart<T>(T? _srcPart, Func<T?> _fnCreatePart)
            where T : OpenXmlPart
        {
            if (null != _srcPart)
            {
                var destPart = _fnCreatePart();
                if (null != destPart)
                {
                    using (var srcStream = _srcPart.GetStream())
                    {
                        destPart.FeedData(srcStream);
                        return destPart;
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Table storing the pairs of the object and the other one assigned to it.
        /// </summary>
        internal static readonly ConditionalWeakTable<object, WeakReference<object>> AssignedObjectMap = new ConditionalWeakTable<object, WeakReference<object>>();

        /// <summary>
        /// Get the object assign to the special source object
        /// </summary>
        /// <typeparam name="T">Type of the destination object</typeparam>
        /// <param name="_src">The source object which is assigned to the result object</param>
        /// <param name="_creator">A function for creating a new object assigned to the given source object if there is not existed one.</param>
        /// <returns>The object assigned to the special source object or null</returns>
        public static T? SingletonAssignObject<T>(object _src, Func<object, object?>? _creator)
            where T : class
        {
            return SingletonAssignObject(_src, _creator, typeof(T)) as T;
        }

        /// <summary>
        /// <param name="_src">The source object which is assigned to the result object</param>
        /// <param name="_creator">A function for creating a new object assigned to the given source object if there is not existed one.</param>
        /// <param name="_targetType">The type of the destination object</param>
        /// <returns>The object assigned to the special source object or null</returns>
        /// </summary>
        public static object? SingletonAssignObject(object _src, Func<object, object?>? _creator, Type _targetType)
        {
            if (AssignedObjectMap.TryGetValue(_src, out WeakReference<object>? _target))
            {
                if (_target?.TryGetTarget(out object? _value) ?? false)
                {
                    if ((null != _value) && _targetType.IsInstanceOfType(_value))
                    {
                        return _value;
                    }
                }
            }

            try
            {
                object? newInst = _creator?.Invoke(_src);
                if (null != newInst)
                {
                    AssignedObjectMap.AddOrUpdate(_src, new WeakReference<object>(newInst));
                }
                return newInst;
            }
            catch (Exception err)
            {
                Trace.TraceWarning(err.ToString());
                return null;
            }
        }
    }

    /// <summary>
    /// The class for modifiable cellection 
    /// </summary>
    /// <typeparam name="TContainer">The type of the container element</typeparam>
    /// <typeparam name="TElement">The type of the children element</typeparam>
    /// <typeparam name="TBoxedElement">The type of the children element in the collection, which box the TElement</typeparam>
    public abstract class FelisModifiableCollection<TContainer, TElement, TBoxedElement> : IEnumerable<TBoxedElement>
        where TContainer : OpenXmlCompositeElement
        where TElement : OpenXmlElement
        where TBoxedElement : class
    {
        /// <summary>
        /// The container element of the items in the collection
        /// </summary>
        protected readonly TContainer ContainerElement;
        /// <summary>
        /// The cache of the collection
        /// </summary>
        protected Lazy<TBoxedElement[]> collection;

        internal FelisModifiableCollection(TContainer _container)
        {
            ContainerElement = _container;
            collection = new Lazy<TBoxedElement[]>(Array.Empty<TBoxedElement>());
            Reload();
        }

        /// <summary>
        /// Reload the cache of the collection
        /// </summary>
        public void Reload()
        {
            collection = new Lazy<TBoxedElement[]>(() => GetElements().Select(e => BoxingElement(e)).ToArray(), LazyThreadSafetyMode.PublicationOnly);
        }

        /// <summary>
        /// Get the item by the index
        /// </summary>
        /// <param name="_index">The index of the target item</param>
        /// <returns></returns>
        public TBoxedElement? this[int _index] => (_index < 0 || _index >= collection.Value.Length) ? null : collection.Value[_index];

        /// <summary>
        /// Add a new item into the collection
        /// </summary>
        /// <param name="_position">The position the new item will be inserted at. If the value is less than zero, the postion is located from the end of the collection, and -1 means the last one in the collection.</param>
        /// <returns></returns>
        public TBoxedElement? Add(int _position = -1)
        {
            int realPos;
            int oriCount = collection.Value.Length;
            if (oriCount == 0)
            {
                PlaceFirstElement(CreateElement(0));
                realPos = 0;
            }
            else
            {
                realPos = (_position < 0) ? Math.Max(oriCount + _position + 1, 0) : Math.Min(_position, oriCount);
                if (realPos < oriCount)
                {
                    var refRow = collection.Value[realPos];
                    ContainerElement.InsertBefore(CreateElement(realPos), UnboxingElement(refRow));
                    ResetIndexOfElements(collection.Value.Skip(realPos), realPos + 1);
                }
                else
                {
                    var refRow = collection.Value[oriCount - 1];
                    ContainerElement.InsertAfter(CreateElement(realPos), UnboxingElement(refRow));
                }
            }
            Reload();
            return collection.Value[realPos];
        }

        /// <summary>
        /// Delete the item in the special index
        /// </summary>
        /// <param name="_position">The index of the item</param>
        public void Delete(int _position)
        {
            int oriCount = collection.Value.Length;
            int realPos = (_position < 0) ? Math.Max(oriCount + _position + 1, 0) : Math.Min(_position, oriCount);
            UnboxingElement(collection.Value[realPos]).Remove();
            ResetIndexOfElements(collection.Value.Skip(realPos + 1), realPos);
            Reload();
        }

        /// <summary>
        /// Clear all the items in the collection
        /// </summary>
        public void Clear()
        {
            foreach (var item in collection.Value)
            {
                UnboxingElement(item).Remove();
            }
            Reload();
        }

        /// <summary>
        /// Get the count of the items in the collection
        /// </summary>
        public int Count => collection.Value.Length;

        /// <summary>
        /// Creating a new open xml element
        /// </summary>
        /// <param name="_index">The position the new element will be located in.</param>
        /// <returns></returns>
        protected abstract TElement CreateElement(int _index);
        
        /// <summary>
        /// Boxing an open xml element into an operating object
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected abstract TBoxedElement BoxingElement(TElement _element);
        
        /// <summary>
        /// Unboxing an open xml element from an operating object
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected abstract OpenXmlElement UnboxingElement(TBoxedElement _element);

        /// <summary>
        /// List all the items inside the container element
        /// </summary>
        /// <returns></returns>
        protected virtual IEnumerable<TElement> GetElements()
        {
            return ContainerElement.Elements<TElement>();
        }

        /// <summary>
        /// Reset the index of the items after the collection is changed
        /// </summary>
        /// <param name="_list">The list of the items which should be changed</param>
        /// <param name="_startIndex">The index of the first item in the list</param>
        protected virtual void ResetIndexOfElements(IEnumerable<TBoxedElement> _list, int _startIndex)
        { 
        }

        /// <summary>
        /// Add the first element into the container
        /// </summary>
        /// <param name="_newElement"></param>
        protected virtual void PlaceFirstElement(OpenXmlElement _newElement)
        {
            ContainerElement.AddChild(_newElement, false);
        }

        /// <summary>
        /// The enumerator of the items in the collection
        /// </summary>
        /// <returns></returns>
        public IEnumerator GetEnumerator()
        {
            return collection.Value.GetEnumerator();
        }

        /// <summary>
        /// The enumerator of the items in the collection
        /// </summary>
        /// <returns></returns>
        IEnumerator<TBoxedElement> IEnumerable<TBoxedElement>.GetEnumerator()
        {
            return (collection.Value as IEnumerable<TBoxedElement>)!.GetEnumerator();
        }
    }

    /// <summary>
    /// The extension functions for Open XML SDK
    /// </summary>
    public static class OpenXmlSDKExtension
    {
        private static readonly Type[] argTypesForGetOoxmlCtor = new Type[] { typeof(OpenXmlElement[]) };

        /// <summary>
        /// Get child element inside a special Open XML Element.
        /// If the target element is not existed, a new one will be create and append to the end of tree of the special element
        /// </summary>
        /// <typeparam name="T">The type of the result element</typeparam>
        /// <param name="_parent">The special element containing the target element</param>
        /// <param name="_fnCreateor">The customer function for creating a new target element</param>
        /// <returns></returns>
        public static T? ForceGetChild<T>(this OpenXmlElement _parent, Func<T>? _fnCreateor = null)
            where T : OpenXmlElement, new()
        {
            if (null != _parent)
            {
                var target = _parent.GetFirstChild<T>();
                if (null == target)
                {
                    target = _fnCreateor?.Invoke() ?? new T();
                    if ((null != target) && (null == target.Parent))
                    {
                        if (_parent is OpenXmlCompositeElement cElement)
                        {
                            cElement.AddChild(target, false);
                        }
                        else
                        {
                            _parent.AppendChild(target);
                        }
                    }
                }
                return target;
            }
            return null;
        }

        /// <summary>
        /// Get the first child filtered by the checker
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_fnChecker"></param>
        /// <returns></returns>
        public static OpenXmlElement? GetFirstChild(this OpenXmlElement _element, Func<OpenXmlElement, bool> _fnChecker)
        {
            return _element.ChildElements.FirstOrDefault(_fnChecker);
        }

        /// <summary>
        /// Get the first child matched to the special type
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_type"></param>
        /// <returns></returns>
        public static OpenXmlElement? GetFirstChild(this OpenXmlElement _element, Type _type)
        {
            return _element.ChildElements.FirstOrDefault(e => _type.IsInstanceOfType(e));
        }

        /// <summary>
        /// Get the children filtered by the checker
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_fnChecker"></param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlElement> Children(this OpenXmlElement _element, Func<OpenXmlElement, bool> _fnChecker)
        {
            return _element.ChildElements.Where(_fnChecker);
        }

        /// <summary>
        /// Get the children matched to the special type
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_type"></param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlElement> Children(this OpenXmlElement _element, Type _type)
        {
            return _element.ChildElements.Where(e => _type.IsInstanceOfType(e));
        }

        /// <summary>
        /// Remove all the children filtered by the checker
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_fnChecker"></param>
        public static void RemoveAllChildren(this OpenXmlElement _element, Func<OpenXmlElement, bool> _fnChecker)
        {
            var curChild = _element.FirstChild;
            while (curChild != null)
            {
                var nextChild = curChild.NextSibling();
                if (_fnChecker(curChild))
                {
                    _element.RemoveChild(curChild);
                }
                curChild = nextChild;
            }
        }

        /// <summary>
        /// Remove all the children matched to the special type
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_type"></param>
        public static void RemoveAllChildren(this OpenXmlElement _element, Type _type)
        {
            var curChild = _element.FirstChild;
            while (curChild != null)
            {
                var nextChild = curChild.NextSibling();
                if (_type.IsInstanceOfType(curChild))
                {
                    _element.RemoveChild(curChild);
                }
                curChild = nextChild;
            }
        }

        /// <summary>
        /// Insert the element into a parent. This method will try the AddChild method if the parent is instance of OpenXmlCompositeElement
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="_parent"></param>
        /// <param name="_subElement"></param>
        /// <param name="_pos"></param>
        /// <returns></returns>
        public static T InsertElement<T>(this OpenXmlElement? _parent, T _subElement, int? _pos = null)
            where T : OpenXmlElement
        {
            if (null != _parent)
            {
                if ((null != _subElement.Parent) && (_parent != _subElement.Parent))
                {
                    _subElement.Remove();
                }

                if (_parent is OpenXmlCompositeElement cElement)
                {
                    if (cElement.AddChild(_subElement, false))
                    {
                        return _subElement;
                    }
                }

                if (_pos is null)
                {
                    _parent.AppendChild(_subElement);
                }
                else
                {
                    _parent.InsertAt(_subElement, _pos.Value);
                }
            }
            return _subElement;
        }

        /// <summary>
        /// The singleton instance of the validator
        /// </summary>
        private static OpenXmlValidator GlobalValidator = new OpenXmlValidator();

        /// <summary>
        /// Validate the element
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_containParent"></param>
        /// <returns></returns>
        public static IEnumerable<ValidationErrorInfo> Validate(this OpenXmlElement? _element, bool _containParent)
        {
            if (null == _element)
            {
                return Enumerable.Empty<ValidationErrorInfo>();
            }

            var ret =  GlobalValidator.Validate(_element);
            if (_containParent && (null != _element.Parent))
            {
                ret = _element.Parent.Validate(false).Concat(ret);
            }

            return ret;
        }

        /// <summary>
        /// Validate and try to fix the postion of the special element
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        public static bool ValidateThePosition(this OpenXmlElement? _element)
        {
            if (null == _element)
            {
                return false;
            }

            var parent = _element.Parent;
            if (null == parent)
            {
                return false;
            }

            var filter = (ValidationErrorInfo info) => (info.RelatedNode == _element) && (info.ErrorType == ValidationErrorType.Schema);
            var oriNext = _element.NextSibling();
            OpenXmlElement? refPos;
            do
            {
                if (!GlobalValidator.Validate(parent).Where(filter).Any())
                {
                    return true;
                }
                refPos = _element.PreviousSibling();
                if (null != refPos)
                {
                    _element.Remove();
                    parent.InsertBefore(_element, refPos);
                }
            } while (null != refPos);
            refPos = oriNext;
            while (null != refPos)
            {
                _element.Remove();
                parent.InsertAfter(_element, refPos);
                if (!GlobalValidator.Validate(parent).Where(filter).Any())
                {
                    return true;
                }
                refPos = _element.NextSibling();
            }
            _element.Remove();
            if (null != oriNext)
            {
                parent.InsertBefore(_element, oriNext);
            }
            else
            {
                parent.AppendChild(_element);
            }
            return false;
        }
    }

    /// <summary>
    /// Parser for the index of the column's index in sheet
    /// </summary>
    public static class FelisSheetColumnAddressParser
    {
        /// <summary>
        /// Convert index to address
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        public static string ToAddress(uint _index)
        {
            string ret = string.Empty;
            byte[] buf = new byte[1];
            var encoding = Encoding.ASCII;
            for (; _index > 0; _index /= 26)
            {
                _index -= 1;
                var mod = _index % 26;
                buf[0] = (byte)('A' + mod);
                ret = encoding.GetString(buf) + ret;
            }
            return ret;
        }

        /// <summary>
        /// Convert address to index
        /// </summary>
        /// <param name="_addr"></param>
        /// <returns></returns>
        public static uint ToIndex(string _addr)
        {
            uint index = 0;
            foreach(char c in _addr) 
            {
                index *= 26;
                var mod = (uint)(c - 'A');
                index += mod + 1;
            }
            return index;
        }
    }
}
