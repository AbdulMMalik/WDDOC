using _TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using _TableRowHeight = DocumentFormat.OpenXml.Wordprocessing.TableRowHeight;
using _TableRowProperties = DocumentFormat.OpenXml.Wordprocessing.TableRowProperties;

namespace WDDoc
{
    public class TableRow : BaseObject
    {
        public _TableRow _TableRow;

        private uint Height { get; set; }

        public TableRow(uint height)
        {
            _TableRow = new _TableRow();
            _TableRowProperties _trp = new _TableRowProperties();

            _TableRowHeight _trh = new _TableRowHeight() { Val = height };
            _trp.Append(_trh);

            _TableRow.Append(_trp);
        }

        public void AddCell(decimal width=1500, int span=1)
        {
            TableCell tc = new TableCell(width, span);
            tc.AppendBorders();
            _TableRow.Append(tc._TableCell);
        }

        public void AddCell(TableCell tc)
        {
            this._TableRow.Append(tc._TableCell);
        }
    }
}
