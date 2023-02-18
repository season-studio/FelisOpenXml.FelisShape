using DocumentFormat.OpenXml.Drawing.Diagrams;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The interface of the tree containing the shapes
    /// </summary>
    public interface IFelisShapeTree
    {
        /// <summary>
        /// Check if there is any shape in the tree
        /// </summary>
        public bool HasShapes { get; }
        /// <summary>
        /// Get an iterator of the shapes in the tree
        /// </summary>
        public IEnumerable<FelisShape> Shapes { get; }
    }

    /// <summary>
    /// The collection of the externel method of the IFelisShapeTree
    /// </summary>
    public static class FelisShapeTreeExtension
    {
        /// <summary>
        /// Get a shape have the special ID
        /// </summary>
        /// <param name="_tree">The instance of IFelisShapeTree</param>
        /// <param name="_id">The ID of the target shape</param>
        /// <param name="_deep">True for searching the subtree in the shapes. The default value is false.</param>
        /// <returns>The target shape</returns>
        public static FelisShape? GetShapeById(this IFelisShapeTree? _tree, uint _id, bool _deep = false)
        {
            if (null != _tree)
            {
                bool hasGroup = false;

                foreach (var shape in _tree.Shapes)
                {
                    if (shape.Id == _id)
                    {
                        return shape;
                    }
                    else if (shape is IFelisShapeTree)
                    {
                        hasGroup = true;
                    }
                }

                if (hasGroup && _deep)
                {
                    foreach (var shape in _tree.Shapes)
                    {
                        if (shape is IFelisShapeTree subTree)
                        {
                            var ret = GetShapeById(subTree, _id, _deep);
                            if (null != ret)
                            {
                                return ret;
                            }
                        }
                    }
                }
            }
            return null;
        }

        /// <summary>
        /// Get the shape with special type
        /// </summary>
        /// <typeparam name="T">The type of the target shape</typeparam>
        /// <param name="_tree">The shapes tree</param>
        /// <param name="_deep">True for searching the subtree in the shapes. The default value is false.</param>
        /// <returns></returns>
        public static IEnumerable<T> GetShapes<T>(this IFelisShapeTree? _tree, bool _deep = false)
            where T : FelisShape
        {
            if (null == _tree)
            {
                return Array.Empty<T>();
            }

            bool hasGroup = false;
            IEnumerable<T> topIterator;
            if (typeof(T) == typeof(FelisShape))
            {
                topIterator = _tree.Shapes.Where(e =>
                {
                    if (e is IFelisShapeTree)
                    {
                        hasGroup = true;
                    }
                    return e.GetType() == typeof(FelisShape);
                }).Select(e => (e as T)!);
            }
            else
            {
                topIterator = _tree.Shapes.Where(e =>
                {
                    if (e is IFelisShapeTree)
                    {
                        hasGroup = true;
                    }
                    return e is T;
                }).Select(e => (e as T)!);
            }
            if (hasGroup && _deep)
            {
                var subIterator = _tree.Shapes.Where(e => e is IFelisShapeTree).SelectMany(e => GetShapes<T>(e as IFelisShapeTree, _deep));
                return topIterator.Concat(subIterator);
            }
            else
            {
                return topIterator;
            }
        }

        /// <summary>
        /// Get the first shape with the special type
        /// </summary>
        /// <typeparam name="T">The special type of the target shape</typeparam>
        /// <param name="_tree">The shapes tree</param>
        /// <param name="_deep">True for searching the subtree in the shapes. The default value is false.</param>
        /// <returns></returns>
        public static T? GetFirstShape<T>(this IFelisShapeTree? _tree, bool _deep = false)
            where T : FelisShape
        {
            return GetShapes<T>(_tree, _deep).FirstOrDefault();
        }
    }
}
