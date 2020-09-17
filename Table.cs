using _TableWidth = DocumentFormat.OpenXml.Wordprocessing.TableWidth;
using _TableIndentation = DocumentFormat.OpenXml.Wordprocessing.TableIndentation;
using _TableCellMarginDefault = DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault;
using _TableCellLeftMargin = DocumentFormat.OpenXml.Wordprocessing.TableCellRightMargin;
using _TableCellRightMargin = DocumentFormat.OpenXml.Wordprocessing.TableCellRightMargin;
using _Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using _TableProperties = DocumentFormat.OpenXml.Wordprocessing.TableProperties;
using _TableWidthUnitValues = DocumentFormat.OpenXml.Wordprocessing.TableWidthUnitValues;
using _TableWidthValues = DocumentFormat.OpenXml.Wordprocessing.TableWidthValues;


namespace WDDoc
{
    public class Table : BaseObject
    {

        public _Table _Table;

        private decimal Width { get; set; }
        private decimal Intendation { get; set; }
        private enum Align
        {
            LEFT,
            MIDDLE,
            RIGHT
        };

        public Table(decimal width=1000, int indentation=0, short marginLeft=0, short marginRight=0)
        {
            _Table = new _Table();
            _TableProperties _tp = new _TableProperties();

            _TableWidth _tw = new _TableWidth() { Width = width.ToString(), Type = _TableWidthUnitValues.Dxa };
            _tp.Append(_tw);

            _TableIndentation _ti = new _TableIndentation() { Width = indentation, Type = _TableWidthUnitValues.Dxa };
            _tp.Append(_ti);

            _TableCellMarginDefault _md = new _TableCellMarginDefault();
            _TableCellLeftMargin _lm = new _TableCellLeftMargin() { Width = marginLeft, Type = _TableWidthValues.Dxa };
            _TableCellRightMargin _rm = new _TableCellRightMargin() { Width = marginRight, Type = _TableWidthValues.Dxa };
            _md.Append(_lm);
            _md.Append(_rm);

            _Table.Append(_tp);
        }
        
        public void AddRow(uint height)
        {
            TableRow tr = new TableRow(height);
            _Table.Append(tr._TableRow);
        }

        public void AddRow(TableRow tr)
        {
            _Table.Append(tr._TableRow);
        }
    }
}
