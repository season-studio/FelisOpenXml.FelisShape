using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using FelisOpenXml.FelisShape.Base;
using FelisOpenXml.FelisShape.Text;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The table shape
    /// </summary>
    [FelisGraphicFrame(@"http://schemas.openxmlformats.org/drawingml/2006/table")]
    public class FelisTable : FelisGraphicFrame
    {
        /// <summary>
        /// The element of the table data
        /// </summary>
        public readonly A.Table TableElement;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_element"></param>
        protected FelisTable(P.GraphicFrame _element)
            : base(_element)
        {
            var table = ForceGraphicData.GetFirstChild<A.Table>();
            if (null == table)
            {
                ForceGraphicData.AddChild(table = new A.Table(), false);
            }
            TableElement = table;
            TableProperties = new Lazy<A.TableProperties>(() => GetTableInnerProperties<A.TableProperties>(TableElement));
            RowsCollection = new FelisTableRowCollection(TableElement);
        }

        /// <summary>
        /// The properties of the table
        /// </summary>
        protected readonly Lazy<A.TableProperties> TableProperties;

        /// <summary>
        /// The collection of the rows
        /// </summary>
        protected readonly FelisTableRowCollection RowsCollection;

        /// <summary>
        /// Get the rows in the table
        /// </summary>
        public FelisTableRowCollection Rows => RowsCollection;

        /// <summary>
        /// Get the count of the columns in the table
        /// </summary>
        public int ColumnsCount
        {
            get
            {
                var row = Rows[0];
                return (null != row) ? row.Cells.Count : 0;
            }
        }

        /// <summary>
        /// Get a special cell by the coordinate
        /// </summary>
        /// <param name="_row">The row index of the cell</param>
        /// <param name="_col">The column index of the cell</param>
        /// <returns></returns>
        public FelisTableCell? Cell(int _row, int _col)
        {
            return Rows[_row]?.Cells[_col];
        }

        /// <summary>
        /// The marker of the band row
        /// </summary>
        public bool MarkBandRow
        {
            get => TableProperties.Value.BandRow ?? false;
            set => TableProperties.Value.BandRow = value;
        }

        /// <summary>
        /// The marker of the band column
        /// </summary>
        public bool MarkBandColumn
        {
            get => TableProperties.Value.BandColumn ?? false;
            set => TableProperties.Value.BandColumn = value;
        }

        /// <summary>
        /// The marker of the first row
        /// </summary>
        public bool MarkFirstRow
        {
            get => TableProperties.Value.FirstRow ?? false;
            set => TableProperties.Value.FirstRow = value;
        }

        /// <summary>
        /// The marker of the first column
        /// </summary>
        public bool MarkFirstColumn
        {
            get => TableProperties.Value.FirstColumn ?? false;
            set => TableProperties.Value.FirstColumn = value;
        }

        /// <summary>
        /// The marker of the last row
        /// </summary>
        public bool MarkLastRow
        {
            get => TableProperties.Value.LastRow ?? false;
            set => TableProperties.Value.LastRow = value;
        }

        /// <summary>
        /// The marker of th last column
        /// </summary>
        public bool MarkLastColumn
        {
            get => TableProperties.Value.LastColumn ?? false;
            set => TableProperties.Value.LastColumn = value;
        }

        /// <summary>
        /// Indicate if the table is right-to-left style
        /// </summary>
        public bool RightToLeft
        {
            get => TableProperties.Value.RightToLeft ?? false;
            set => TableProperties.Value.RightToLeft = value;
        }

        /// <summary>
        /// Get the properties element of the table's content
        /// </summary>
        internal static T GetTableInnerProperties<T>(OpenXmlElement _element)
            where T : OpenXmlElement, new()
        {
            T? props = _element.GetFirstChild<T>();
            if (null == props)
            {
                props = new T();
                _element.InsertElement(props);
            }
            return props;
        }
    }

    /// <summary>
    /// The table row class
    /// </summary>
    public class FelisTableRow : FelisCompositeElement
    {
        internal FelisTableRow(A.TableRow _element)
            : base(_element) 
        {
            cellsCollection = new FelisTableCellCollection((Element as A.TableRow)!);
        }

        /// <summary>
        /// the cache of the collection of the cell
        /// </summary>
        protected readonly FelisTableCellCollection cellsCollection;

        /// <summary>
        /// Get the cells in the row
        /// </summary>
        public FelisTableCellCollection Cells => cellsCollection;
    }

    /// <summary>
    /// The collection of the row
    /// </summary>
    public class FelisTableRowCollection : FelisModifiableCollection<A.Table, A.TableRow, FelisTableRow>
    {
        internal FelisTableRowCollection(A.Table _table)
            : base(_table)
        {
            
        }

        /// <summary>
        /// Create an empty row element
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        protected override A.TableRow CreateElement(int _index)
        {
            return new A.TableRow();
        }

        /// <summary>
        /// Boxing the row element
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override FelisTableRow BoxingElement(A.TableRow _element)
        {
            return new FelisTableRow(_element);
        }

        /// <summary>
        /// Unboxing the row object
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override OpenXmlElement UnboxingElement(FelisTableRow _element)
        {
            return _element.Element;
        }
    }

    /// <summary>
    /// The collection of the cell
    /// </summary>
    public class FelisTableCellCollection : FelisModifiableCollection<A.TableRow, A.TableCell, FelisTableCell>
    {
        internal FelisTableCellCollection(A.TableRow _tableRow)
            : base(_tableRow)
        {
        }

        /// <summary>
        /// Create an empty cell element
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        protected override A.TableCell CreateElement(int _index)
        {
            return new A.TableCell();
        }

        /// <summary>
        /// Boxing the cell element
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override FelisTableCell BoxingElement(A.TableCell _element)
        {
            return new FelisTableCell(_element);
        }

        /// <summary>
        /// Unboxing the cell object
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override OpenXmlElement UnboxingElement(FelisTableCell _element)
        {
            return _element.Element;
        }
    }

    /// <summary>
    /// The table cell class
    /// </summary>
    public class FelisTableCell : FelisCompositeElement
    {
        internal FelisTableCell(A.TableCell _element)
            : base(_element) 
        { 
        }

        /// <summary>
        /// Check if the cell contains text content
        /// </summary>
        public bool HasTextContent
        {
            get
            {
                return Element.GetFirstChild<A.TextBody>()?.Descendants<A.Text>().FirstOrDefault((text) => !string.IsNullOrEmpty(text.Text)) != null;
            }
        }

        /// <summary>
        /// Get the text body in the cell
        /// </summary>
        public FelisTextBody? TextBody
        {
            get
            {
                var textBodyElement = Element.GetFirstChild<A.TextBody>();
                return (null == textBodyElement) ? null : new FelisTextBody(textBodyElement);
            }
        }

        /// <summary>
        /// The count of the rows this cell merges
        /// </summary>
        public int MergeRows
        {
            get => (Element as A.TableCell)!.RowSpan ?? 0;
            set => (Element as A.TableCell)!.RowSpan = value;
        }

        /// <summary>
        /// The count of the columns this cell merges
        /// </summary>
        public int MergeColumns
        {
            get => (Element as A.TableCell)!.GridSpan ?? 0;
            set => (Element as A.TableCell)!.GridSpan = value;
        }

        /// <summary>
        /// Check if this cell is the master of the merged cells
        /// </summary>
        public bool IsMergeMaster => (MergeRows > 0) || (MergeColumns > 0);

        /// <summary>
        /// Check if this cell is merged by others in row
        /// </summary>
        public bool IsMergedByRow
        {
            get => (Element as A.TableCell)!.VerticalMerge ?? false;
            set => (Element as A.TableCell)!.VerticalMerge = value;
        }

        /// <summary>
        /// Check if this cell is merged by others in column
        /// </summary>
        public bool IsMergedByColumn
        {
            get => (Element as A.TableCell)!.HorizontalMerge ?? false;
            set => (Element as A.TableCell)!.HorizontalMerge = value;
        }

        /// <summary>
        /// Check if this cell is merged by others
        /// </summary>
        public bool IsMerged => IsMergedByRow || IsMergedByColumn;
    }
}
