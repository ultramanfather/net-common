using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace OpenXmlConverter
{
    using a = DocumentFormat.OpenXml.Drawing;
    using pic = DocumentFormat.OpenXml.Drawing.Pictures;
    using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
    using wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
    using mc = DocumentFormat.OpenXml;
    using v = DocumentFormat.OpenXml.Vml;
    using w10 = DocumentFormat.OpenXml.Vml.Wordprocessing;

    /// <summary>
    /// 定位点
    /// </summary>
    struct AnchorPoint
    {
        /// <summary>
        /// shape的宽度
        /// </summary>
        private Unit Width { get; set; }

        /// <summary>
        /// shape的高度
        /// </summary>
        private Unit Height { get; set; }

        /// <summary>
        /// 距离顶部的长度
        /// </summary>
        private Unit Top { get; set; }

        /// <summary>
        /// 距离左边的长度
        /// </summary>
        private Unit Left { get; set; }

        /// <summary>
        /// 文档属性ID
        /// </summary>
        private UInt32Value DocPrId { get; set; }

        public AnchorPoint(int id, Unit width, Unit height, Unit left, Unit top)
        {
            DocPrId = UInt32Value.FromUInt32((uint) id);
            Width = width;
            Height = height;
            Top = top;
            Left = left;
        }

        #region AddImagePart

        /// <summary>
        /// 添加定位点
        /// </summary>
        /// <returns></returns>
        public OpenXmlElement AddAnchorPoint()
        {
            string style =
                $"position:absolute;margin-left:{Left.ValueInPoint}pt;margin-top:{Top.ValueInPoint}pt;width:{Width.ValueInPoint}pt;height:{Height.ValueInPoint}pt;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:0;mso-wrap-distance-top:0;mso-wrap-distance-right:0;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:page;mso-position-vertical:absolute;mso-position-vertical-relative:page;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:middle";

            return new Picture(
                new v.Rectangle(new w10.TextWrap
                    {
                        AnchorX = w10.HorizontalAnchorValues.Page,
                        AnchorY = w10.VerticalAnchorValues.Page,
                    }
                )
                {
                    FillColor = "black [3213]",
                    Style = style,
                    Gfxdata =
                        "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF&#xA;90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA&#xA;0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD&#xA;OlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893&#xA;SUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y&#xA;JsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl&#xA;bHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR&#xA;JVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY&#xA;22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i&#xA;OWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA&#xA;IQASZvZulwIAAIAFAAAOAAAAZHJzL2Uyb0RvYy54bWysVM1u2zAMvg/YOwi6r46ztuuCOkXQosOA&#xA;og3WDj2rshQLkERNUuJkLzNgtz7EHmfYa4ySHSftih2G+SCTIvnxRyRPz9ZGk5XwQYGtaHkwokRY&#xA;DrWyi4p+vrt8c0JJiMzWTIMVFd2IQM+mr1+dtm4ixtCAroUnCGLDpHUVbWJ0k6IIvBGGhQNwwqJQ&#xA;gjcsIusXRe1Zi+hGF+PR6LhowdfOAxch4O1FJ6TTjC+l4PFGyiAi0RXF2GI+fT4f0llMT9lk4Zlr&#xA;FO/DYP8QhWHKotMB6oJFRpZe/QFlFPcQQMYDDqYAKRUXOQfMphw9y+a2YU7kXLA4wQ1lCv8Pll+v&#xA;5p6ouqJjSiwz+ES/vj3+/PGdjFNtWhcmqHLr5r7nApIp0bX0Jv0xBbLO9dwM9RTrSDhevj0e4UcJ&#xA;R1F5kmlEKXbGzof4QYAhiaiox+fKVWSrqxA71a1K8hVAq/pSaZ2Z1CLiXHuyYvi4cV2mgBH8iZa2&#xA;SddCsurE6aZIeXWZZCputEh62n4SEquBsY9zILkPd04Y58LGshM1rBad76O91AaLHEsGTMgS/Q/Y&#xA;PcDTBLbYXZS9fjIVuY0H49HfAuuMB4vsGWwcjI2y4F8C0JhV77nT3xapK02q0gPUG+wVD90QBccv&#xA;FT7bFQtxzjxODb40boJ4g4fU0FYUeoqSBvzXl+6TPjYzSilpcQorGr4smReU6I8W2/x9eXiYxjYz&#xA;h0fvxsj4fcnDvsQuzTlgL5S4cxzPZNKPektKD+YeF8YseUURsxx9V5RHv2XOY7cdcOVwMZtlNRxV&#xA;x+KVvXU8gaeqpra8W98z7/rejdj017CdWDZ51sKdbrK0MFtGkCr3966ufb1xzHPj9Csp7ZF9Pmvt&#xA;Fuf0NwAAAP//AwBQSwMEFAAGAAgAAAAhAKkYkKvgAAAACQEAAA8AAABkcnMvZG93bnJldi54bWxM&#xA;j8tOwzAQRfdI/IM1SGxQ69TqI4Q4VQvtphug7Qe48ZBEjcdR7Lbh7xlWsJyZozvn5svBteKKfWg8&#xA;aZiMExBIpbcNVRqOh+0oBRGiIWtaT6jhGwMsi/u73GTW3+gTr/tYCQ6hkBkNdYxdJmUoa3QmjH2H&#xA;xLcv3zsTeewraXtz43DXSpUkc+lMQ/yhNh2+1lie9xen4fD+Md9smxmp7vy2mpaz9dNmt9b68WFY&#xA;vYCIOMQ/GH71WR0Kdjr5C9kgWg2pUopRDSOVTkEwkS7SZxAn3kxAFrn836D4AQAA//8DAFBLAQIt&#xA;ABQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAAAAAAAAAAAAAAAAAAABbQ29udGVudF9UeXBlc10u&#xA;eG1sUEsBAi0AFAAGAAgAAAAhADj9If/WAAAAlAEAAAsAAAAAAAAAAAAAAAAALwEAAF9yZWxzLy5y&#xA;ZWxzUEsBAi0AFAAGAAgAAAAhABJm9m6XAgAAgAUAAA4AAAAAAAAAAAAAAAAALgIAAGRycy9lMm9E&#xA;b2MueG1sUEsBAi0AFAAGAAgAAAAhAKkYkKvgAAAACQEAAA8AAAAAAAAAAAAAAAAA8QQAAGRycy9k&#xA;b3ducmV2LnhtbFBLBQYAAAAABAAEAPMAAAD+BQAAAAA=&#xA;",
                    Stroked = false,
                    Id = "矩形 " + DocPrId,
                }
            );
        }

        #endregion
    }
}