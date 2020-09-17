using Wp = DocumentFormat.OpenXml.Wordprocessing;
using DWp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using OXML = DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WDDoc
{
    public class Drawing : BaseObject
    {
        public class Picture
        {
            public enum ALIGNMENT
            {
                ACNHOR,
                INLINE
            }

            public static Wp.Drawing GetAnchorPicture(String imagePartId, uint width = 1500, uint height = 1500, uint horizontalOffset = 0, uint verticalOffset = 0, String pictureName = "Picture")
            {
                Wp.Drawing _drawing = new Wp.Drawing();
                DWp.Anchor _anchor = new DWp.Anchor()
                {
                    DistanceFromTop = (OXML.UInt32Value)0U,
                    DistanceFromBottom = (OXML.UInt32Value)0U,
                    DistanceFromLeft = (OXML.UInt32Value)0U,
                    DistanceFromRight = (OXML.UInt32Value)0U,
                    SimplePos = false,
                    RelativeHeight = (OXML.UInt32Value)0U,
                    BehindDoc = true,
                    Locked = false,
                    LayoutInCell = true,
                    AllowOverlap = true,
                    EditId = "44CEF5E4",
                    AnchorId = "44803ED1"
                };
                DWp.SimplePosition _spos = new DWp.SimplePosition()
                {
                    X = 0L,
                    Y = 0L
                };

                DWp.HorizontalPosition _hp = new DWp.HorizontalPosition()
                {
                    RelativeFrom = DWp.HorizontalRelativePositionValues.Column
                };
                DWp.PositionOffset _hPO = new DWp.PositionOffset();
                _hPO.Text = horizontalOffset.ToString();
                _hp.Append(_hPO);

                DWp.VerticalPosition _vp = new DWp.VerticalPosition()
                {
                    RelativeFrom = DWp.VerticalRelativePositionValues.Paragraph
                };
                DWp.PositionOffset _vPO = new DWp.PositionOffset();
                _vPO.Text = verticalOffset.ToString();
                _vp.Append(_vPO);

                DWp.Extent _e = new DWp.Extent()
                {
                    Cx = height,
                    Cy = width
                };

                DWp.EffectExtent _ee = new DWp.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                };

                DWp.WrapTight _wp = new DWp.WrapTight()
                {
                    WrapText = DWp.WrapTextValues.BothSides
                };

                DWp.WrapPolygon _wpp = new DWp.WrapPolygon()
                {
                    Edited = false
                };
                DWp.StartPoint _sp = new DWp.StartPoint()
                {
                    X = 0L,
                    Y = 0L
                };

                DWp.LineTo _l1 = new DWp.LineTo() { X = 0L, Y = 0L };
                DWp.LineTo _l2 = new DWp.LineTo() { X = 0L, Y = 0L };
                DWp.LineTo _l3 = new DWp.LineTo() { X = 0L, Y = 0L };
                DWp.LineTo _l4 = new DWp.LineTo() { X = 0L, Y = 0L };

                _wpp.Append(_sp);
                _wpp.Append(_l1);
                _wpp.Append(_l2);
                _wpp.Append(_l3);
                _wpp.Append(_l4);

                _wp.Append(_wpp);

                DWp.DocProperties _dp = new DWp.DocProperties()
                {
                    Id = 1U,
                    Name = pictureName
                };

                OXML.Drawing.Graphic _g = new OXML.Drawing.Graphic();
                OXML.Drawing.GraphicData _gd = new OXML.Drawing.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };
                OXML.Drawing.Pictures.Picture _pic = new OXML.Drawing.Pictures.Picture();

                OXML.Drawing.Pictures.NonVisualPictureProperties _nvpp = new OXML.Drawing.Pictures.NonVisualPictureProperties();
                OXML.Drawing.Pictures.NonVisualDrawingProperties _nvdp = new OXML.Drawing.Pictures.NonVisualDrawingProperties()
                {
                    Id = 0,
                    Name = pictureName
                };
                OXML.Drawing.Pictures.NonVisualPictureDrawingProperties _nvpdp = new OXML.Drawing.Pictures.NonVisualPictureDrawingProperties();
                _nvpp.Append(_nvdp);
                _nvpp.Append(_nvpdp);


                OXML.Drawing.Pictures.BlipFill _bf = new OXML.Drawing.Pictures.BlipFill();
                OXML.Drawing.Blip _b = new OXML.Drawing.Blip()
                {
                    Embed = imagePartId,
                    CompressionState = OXML.Drawing.BlipCompressionValues.Print
                };
                _bf.Append(_b);

                OXML.Drawing.Stretch _str = new OXML.Drawing.Stretch();
                OXML.Drawing.FillRectangle _fr = new OXML.Drawing.FillRectangle();
                _str.Append(_fr);
                _bf.Append(_str);

                OXML.Drawing.Pictures.ShapeProperties _shp = new OXML.Drawing.Pictures.ShapeProperties();
                OXML.Drawing.Transform2D _t2d = new OXML.Drawing.Transform2D();
                OXML.Drawing.Offset _os = new OXML.Drawing.Offset()
                {
                    X = 0L,
                    Y = 0L
                };
                OXML.Drawing.Extents _ex = new OXML.Drawing.Extents()
                {
                    Cx = 989965L,
                    Cy = 791845L
                };

                _t2d.Append(_os);
                _t2d.Append(_ex);

                OXML.Drawing.PresetGeometry _preGeo = new OXML.Drawing.PresetGeometry()
                {
                    Preset = OXML.Drawing.ShapeTypeValues.Rectangle
                };
                OXML.Drawing.AdjustValueList _adl = new OXML.Drawing.AdjustValueList();
                _preGeo.Append(_adl);

                _shp.Append(_t2d);
                _shp.Append(_preGeo);

                _pic.Append(_nvpp);
                _pic.Append(_bf);
                _pic.Append(_shp);

                _gd.Append(_pic);
                _g.Append(_gd);

                _anchor.Append(_spos);
                _anchor.Append(_hp);
                _anchor.Append(_vp);
                _anchor.Append(_e);
                _anchor.Append(_ee);
                _anchor.Append(_wp);
                _anchor.Append(_dp);
                _anchor.Append(_g);

                _drawing.Append(_anchor);

                return _drawing;
            }
        }
    }
}
