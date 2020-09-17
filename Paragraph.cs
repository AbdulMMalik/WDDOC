using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OXML = DocumentFormat.OpenXml;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

namespace WDDoc
{
    public class Paragraph : BaseObject
    {
        public Wp.Paragraph _Paragraph;

        public enum JUSTIFICATION
        {
            LEFT,
            CENTER,
            RIGHT,
            NONE
        };

        public Paragraph(uint leftIntendation=0, uint rightIntendation=0, JUSTIFICATION justification=JUSTIFICATION.NONE)
        {
            _Paragraph = new Wp.Paragraph()
            {
                RsidParagraphMarkRevision = "00297913",
                RsidParagraphAddition = "00297913",
                RsidParagraphProperties = "00297913",
                RsidRunAdditionDefault = "0074062A",
                ParagraphId = "29551F12",
                TextId = "11181D94"
            };

            Wp.ParagraphProperties _pp = new Wp.ParagraphProperties();
            Wp.Indentation _i = new Wp.Indentation()
            {
                Left = leftIntendation.ToString(),
                Right = rightIntendation.ToString()
            };
            _pp.Append(_i);

            if(justification.Equals(JUSTIFICATION.NONE) is false)
            {
                Wp.Justification _j;
                if (justification.Equals(JUSTIFICATION.CENTER) is true)
                    _j = new Wp.Justification() { Val = Wp.JustificationValues.Center };
                else if (justification.Equals(JUSTIFICATION.LEFT) is true)
                    _j = new Wp.Justification() { Val = Wp.JustificationValues.Left };
                else
                    _j = new Wp.Justification() { Val = Wp.JustificationValues.Right };

                _pp.Append(_j);
            }

            _Paragraph.Append(_pp);
        }

        public void AddText(String text="", uint fontSize=12, bool bold=false, bool italic=false)
        {
            Wp.Run _r = new Wp.Run();
            Wp.RunProperties _rp = new Wp.RunProperties();
            Wp.RunFonts _rf = new Wp.RunFonts()
            {
                Ascii = "Calibri",
                HighAnsi = "Calibri",
                EastAsia = "Calibri",
                ComplexScript = "Calibri"
            };
            Wp.FontSize _fs = new Wp.FontSize()
            {
                Val = fontSize.ToString()
            };
            Wp.Languages _l = new Wp.Languages()
            {
                Bidi = "en-us"
            };

            _rp.Append(_rf);
            _rp.Append(_fs);
            _rp.Append(_l);

            if (italic)
                _rp.Append(new Wp.Italic());

            if (bold)
                _rp.Append(new Wp.Bold());

            Wp.Text _t = new Wp.Text();
            _t.Text = text;

            _r.Append(_rp);
            _r.Append(_t);

            this._Paragraph.Append(_r);
        }
    }
}
