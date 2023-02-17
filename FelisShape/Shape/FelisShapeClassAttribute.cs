using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;
using System.Linq.Expressions;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The description attribute of a shape class
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    internal class FelisShapeClassAttribute : Attribute
    {
        internal delegate object CreateHandler(OpenXmlCompositeElement _element);
        internal static Type[] CreatorParameterTypes = new[] { typeof(OpenXmlCompositeElement) };

        public readonly Type ShapeType;
        public Type[]? NonVisualDrawingPropertiesChain = null;
        public CreateHandler? Creator { get; private set; } = null;
        public Type? FactoryType { get; private set; } = null;

        public FelisShapeClassAttribute(Type ShapeType)
        {
            this.ShapeType = ShapeType;
        }

        public void Assign(FelisShapeClassAttribute _other)
        {
            if (null == NonVisualDrawingPropertiesChain)
            {
                NonVisualDrawingPropertiesChain = _other.NonVisualDrawingPropertiesChain;
            }
        }

        public static CreateHandler GenerateCreatorDelegate(MethodBase _callInfo)
        {
            var arg = Expression.Parameter(typeof(OpenXmlCompositeElement), "arg");
            Expression argAs = (_callInfo.GetParameters()[0].ParameterType == typeof(OpenXmlCompositeElement))
                                    ? arg : Expression.TypeAs(arg, _callInfo.GetParameters()[0].ParameterType);
            Expression call;
            if (_callInfo is MethodInfo method)
            {
                call = Expression.Call(null, method, argAs);
            }
            else if (_callInfo is ConstructorInfo ctor)
            {
                call = Expression.New(ctor, argAs);
            }
            else
            {
                throw new ArgumentException("_callInfo must be a method or a constructor");
            }
            var testNull = Expression.Equal(argAs, Expression.Constant(null));
            var root = Expression.Condition(testNull, Expression.Default(typeof(object)), Expression.TypeAs(call, typeof(object)));
            var fn = Expression.Lambda<CreateHandler>(root, arg);
            return fn.Compile();
        }

        public bool TrySetCreator(Type _mapType)
        {
            CreateHandler? creator = null;
            var specialsCtorTypes = new[] { ShapeType };
            var createMethodInfo = _mapType.GetMethod("FromElement", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, specialsCtorTypes);
            if ((null == createMethodInfo) || ((createMethodInfo.ReturnType != typeof(FelisShape)) && !createMethodInfo.ReturnType.IsSubclassOf(typeof(FelisShape))))
            {
                createMethodInfo = _mapType.GetMethod("FromElement", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, CreatorParameterTypes);
            }
            if ((null != createMethodInfo) && ((createMethodInfo.ReturnType == typeof(FelisShape)) || createMethodInfo.ReturnType.IsSubclassOf(typeof(FelisShape))))
            {
                creator = GenerateCreatorDelegate(createMethodInfo);// createMethodInfo.CreateDelegate<CreateHandler>();//(OpenXmlCompositeElement _element) => createMethodInfo.Invoke(null, new[] { _element })!;
            }
            if ((null == creator) && (!_mapType.IsAbstract) && (!_mapType.IsGenericType) && (!_mapType.IsInterface))
            {
                var ctorInfo = _mapType.GetConstructor(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.CreateInstance | BindingFlags.Instance, specialsCtorTypes);
                if (null == ctorInfo)
                {
                    ctorInfo = _mapType.GetConstructor(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.CreateInstance | BindingFlags.Instance, CreatorParameterTypes);
                }
                if (null != ctorInfo)
                {
                    creator = GenerateCreatorDelegate(ctorInfo);// (OpenXmlCompositeElement _element) => ctorInfo.Invoke(new[] { _element });
                }
            }
            if (null != creator)
            {
                Creator = creator;
                FactoryType = _mapType;
                return true;
            }
            return false;
        }

        public P.NonVisualDrawingProperties? GetNonVisualDrawingProperties(OpenXmlCompositeElement _element)
        {
            if (null != NonVisualDrawingPropertiesChain)
            {
                OpenXmlElement? parent = _element;
                foreach (var item in NonVisualDrawingPropertiesChain)
                {
                    parent = _element.ChildElements.SingleOrDefault((e) => item.IsInstanceOfType(e));
                    if (null == parent)
                    {
                        break;
                    }
                }
                return (null == parent) ? null : parent.GetFirstChild<P.NonVisualDrawingProperties>();
            }
            return null;
        }

        public OpenXmlElement? GetNonVisualProperties(OpenXmlCompositeElement _element)
        {
            if (null != NonVisualDrawingPropertiesChain)
            {
                OpenXmlElement? target = _element;
                foreach (var item in NonVisualDrawingPropertiesChain)
                {
                    target = _element.ChildElements.SingleOrDefault((e) => item.IsInstanceOfType(e));
                    if (null == target)
                    {
                        break;
                    }
                }
                return target;
            }
            return null;
        }
    }
}
