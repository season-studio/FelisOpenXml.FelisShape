using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2010.CustomUI;
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
            if (null != _tree)
            {
                IEnumerable<IFelisShapeTree> groupSet = Array.Empty<IFelisShapeTree>();
                foreach (var shape in _tree.Shapes)
                {
                    if (shape is IFelisShapeTree subTree)
                    {
                        if (_deep)
                        {
                            groupSet = groupSet.Append(subTree);
                        }
                    }
                    else if (typeof(T) == typeof(FelisShape))
                    {
                        if (shape.GetType() == typeof(FelisShape))
                        {
                            yield return (shape as T)!;
                        }
                    }
                    else if (shape is T shapeT)
                    {
                        yield return shapeT;
                    }
                }

                if (_deep && groupSet.Any())
                {
                    foreach (var subTree in groupSet)
                    {
                        foreach (var shape in GetShapes<T>(subTree, _deep))
                        {
                            yield return shape;
                        }
                    }
                }
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
            var r = GetShapes<T>(_tree, _deep);
            return r.FirstOrDefault();
        }
    }
}
