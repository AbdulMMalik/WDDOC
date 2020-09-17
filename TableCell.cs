using _TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using _TableCellWidth = DocumentFormat.OpenXml.Wordprocessing.TableCellWidth;
using _TableCellBorders = DocumentFormat.OpenXml.Wordprocessing.TableCellBorders;
using _TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using _LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using _RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using _BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using _BorderValues = DocumentFormat.OpenXml.Wordprocessing.BorderValues;
using _TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using _TableWidthUnitValues = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues;
using _GridSpan = DocumentFormat.OpenXml.Wordprocessing.GridSpan;

namespace WDDoc
{
    public class TableCell : BaseObject
    {
        public _TableCell _TableCell { get; }
        private _TableCellProperties _tcp;

        public TableCell(decimal width=1500, int span=1)
        {
            _TableCell = new _TableCell();
            _tcp = new _TableCellProperties();

            _TableCellWidth _tcw = new _TableCellWidth() { Width = width.ToString(), Type = _TableWidthUnitValues.Dxa };
            _tcp.Append(_tcw);

            _GridSpan gs = new _GridSpan() { Val = span };
            _tcp.Append(gs);

            _TableCell.Append(_tcp);
        }

        public void AppendBorders(_BorderValues borderValue = _BorderValues.Single, string color = "000000", uint size = 12)
        {
            _TableCellBorders _tcb = new _TableCellBorders();
            _TopBorder _tb = new _TopBorder() { Val = borderValue, Color = color, Size = size, Space = (DocumentFormat.OpenXml.UInt32Value)0U };
            _LeftBorder _lb = new _LeftBorder() { Val = borderValue, Color = color, Size = size, Space = (DocumentFormat.OpenXml.UInt32Value)0U };
            _BottomBorder _bb = new _BottomBorder() { Val = borderValue, Color = color, Size = size, Space = (DocumentFormat.OpenXml.UInt32Value)0U };
            _RightBorder _rb = new _RightBorder() { Val = borderValue, Color = color, Size = size, Space = (DocumentFormat.OpenXml.UInt32Value)0U };

            _tcb.Append(_tb);
            _tcb.Append(_lb);
            _tcb.Append(_bb);
            _tcb.Append(_rb);

            _tcp.Append(_tcb);
        }

        public void AppendLeftBorder(_BorderValues borderValue, string color, uint size)
        {

        }

        public void AppendRightBorder(_BorderValues borderValues, string color, uint size)
        {

        }
    }
}
