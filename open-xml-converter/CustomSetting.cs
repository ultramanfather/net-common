using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OpenXmlConverter.IO;
using System;
using System.Collections.Generic;
using System.IO;
using Columns = DocumentFormat.OpenXml.Wordprocessing.Columns;
using Header = DocumentFormat.OpenXml.Wordprocessing.Header;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OpenXmlConverter
{
    using a = DocumentFormat.OpenXml.Drawing;
    using a14 = DocumentFormat.OpenXml.Office2010.Drawing;
    using pic = DocumentFormat.OpenXml.Drawing.Pictures;
    using wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
    using wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;

    public class CustomSetting
    {
        private Library.Model.ExportWordSetting settings;
        private MainDocumentPart mainPart;
        private Unit pageWidth;
        private Unit pageHeight;
        private int headImgId = 1;
        int docPrId = 1000;
        double pointWidth = 0.7d;
        double pointHeight = 0.35d;
        double pointLeft;
        Unit pointWidthUnit;
        Unit pointHeightUnit;

        private Settings docSettings;

        Unit marginTop;
        Unit marginBottom;
        Unit marginLeft;
        Unit marginRight;

        // logo宽度固定
        private Unit logoHeight;


        // QRCode宽高固定
        private Unit qrCodeSize;

        // 是否是横向
        bool isLandscape = false;

        // 内容宽度
        Unit contentWidth;
        // 内容高度
        Unit contentHeight;

        public CustomSetting(Library.Model.ExportWordSetting settings, MainDocumentPart mainPart)
        {
            this.settings = settings;
            this.mainPart = mainPart;
            marginTop = new Unit(UnitMetric.Centimeter, settings.MarginTop);
            marginBottom = new Unit(UnitMetric.Centimeter, settings.MarginBottom);
            marginLeft = new Unit(UnitMetric.Centimeter, settings.MarginLeft);
            marginRight = new Unit(UnitMetric.Centimeter, settings.MarginRight);
            pageWidth = new Unit(UnitMetric.Centimeter, settings.PageWidth);
            pageHeight = new Unit(UnitMetric.Centimeter, settings.PageHeight);
            contentWidth = new Unit(UnitMetric.Centimeter, settings.PageWidth - settings.MarginLeft - settings.MarginRight);
            contentHeight = new Unit(UnitMetric.Centimeter, settings.PageHeight - settings.MarginTop - settings.MarginBottom);
            logoHeight = Unit.Parse("1cm");
            qrCodeSize = Unit.Parse("1.5cm");
        }

        /// <summary>
        /// 页面设置
        /// </summary>
        private void PageSetup()
        {
            SectionProperties sectionProperties = mainPart.Document.Body.GetFirstChild<SectionProperties>();

            sectionProperties?.Remove();
            mainPart.Document.Body.Append(SectionProperties());
        }

        /// <summary>
        /// 页面属性
        /// </summary>
        /// <returns></returns>
        private SectionProperties SectionProperties()
        {

            isLandscape = settings.PageWidth > settings.PageHeight;
            Unit marginHeader = new Unit(UnitMetric.Centimeter, 0.6);
            Unit marginFooter = new Unit(UnitMetric.Centimeter, 0.6);

            return new SectionProperties(
                new PageSize
                {
                    Width = (UInt32Value)pageWidth.ValueInDxa,
                    Height = (UInt32Value)pageHeight.ValueInDxa,
                    Orient = settings.PageWidth < settings.PageHeight
                        ? PageOrientationValues.Portrait
                        : PageOrientationValues.Landscape,
                    Code = null,
                },
                new PageMargin()
                {
                    Top = (Int32Value)marginTop.ValueInDxa,
                    Bottom = (Int32Value)marginBottom.ValueInDxa,
                    Left = (UInt32Value)marginLeft.ValueInDxa,
                    Right = (UInt32Value)marginRight.ValueInDxa,
                    Header = (UInt32Value)marginHeader.ValueInDxa,
                    Footer = (UInt32Value)marginFooter.ValueInDxa,
                    Gutter = 0U
                },
                new MirrorMargins()
                {
                    //0表示不对称,非0表示对称
                    Val = settings.MirrorMargin == 1
                },
                new Columns
                {
                    ColumnCount = (Int16Value)settings.Columns,
                    Space = Convert.ToString(settings.ColumnSpacing * 220),
                    Separator = OnOffValue.FromBoolean(settings.HasLine != 0)
                },
                new DocGrid
                {
                    LinePitch = 303,
                    Type = DocGridValues.Lines
                });
        }

        /// <summary>
        /// 页眉页脚设置
        /// </summary>
        private void HeaderFooter()
        {
            // Delete the existing header and footer parts
            // mainPart.DeleteParts(mainPart.HeaderParts);
            mainPart.DeleteParts(mainPart.FooterParts);

            // Create a new header and footer part
            // HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();

            // Get Id of the headerPart and footer parts
            // string headerPartId = mainPart.GetIdOfPart(headerPart);
            string footerPartId = mainPart.GetIdOfPart(footerPart);

            GenerateHeaderPartContent();
            GenerateFooterPartContent(footerPart);

            // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

            foreach (var section in sections)
            {
                // Delete existing references to headers and footers
                // section.RemoveAllChildren<HeaderReference>();
                section.RemoveAllChildren<FooterReference>();

                // Create the new header and footer reference node
                // section.PrependChild(new HeaderReference() {Id = headerPartId});
                section.PrependChild(new FooterReference() { Id = footerPartId, Type = HeaderFooterValues.Default });
                section.PrependChild(new FooterReference() { Id = footerPartId, Type = HeaderFooterValues.Even });
            }
        }

        /// <summary>
        /// 创建一个新的页眉对象
        /// </summary>
        private Header NewHeader()
        {
            Header header = new Header()
            {
                MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex wp14" }
            };
            header.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header.AddNamespaceDeclaration("cx3", "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex");
            header.AddNamespaceDeclaration("cx4", "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex");
            header.AddNamespaceDeclaration("cx5", "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex");
            header.AddNamespaceDeclaration("cx6", "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex");
            header.AddNamespaceDeclaration("cx7", "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex");
            header.AddNamespaceDeclaration("cx8", "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex");
            header.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header.AddNamespaceDeclaration("aink", "http://schemas.microsoft.com/office/drawing/2016/ink");
            header.AddNamespaceDeclaration("am3d", "http://schemas.microsoft.com/office/drawing/2017/model3d");
            header.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            header.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            header.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            header.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");
            return header;
        }

        /// <summary>
        /// 空页眉对象
        /// </summary>
        /// <returns></returns>
        private Header EmptyHeader()
        {
            Header header = NewHeader();
            header.Append(new Paragraph());
            return header;
        }

        private Settings NewSettings()
        {
            Settings _settings = new Settings()
            { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se w16cid w16 w16cex" } };
            _settings.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            _settings.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            _settings.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            _settings.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            _settings.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            _settings.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            _settings.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            _settings.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            _settings.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            _settings.AddNamespaceDeclaration("w16cex", "http://schemas.microsoft.com/office/word/2018/wordml/cex");
            _settings.AddNamespaceDeclaration("w16cid", "http://schemas.microsoft.com/office/word/2016/wordml/cid");
            _settings.AddNamespaceDeclaration("w16", "http://schemas.microsoft.com/office/word/2018/wordml");
            _settings.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            _settings.AddNamespaceDeclaration("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main");
            return _settings;
        }

        /// <summary>
        /// 机构Logo
        /// </summary>
        /// <returns></returns>
        private Run Logo(HeaderPart headerPart)
        {
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Png);

            Size size;
            string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? string.Empty,
                "Resources",
                (isLandscape ? "landscape" : "portrait") + "_logo.png");
            using (FileStream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                size = ImagePrefetcher.GetImageSize(stream);
                stream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(stream);
            }

            // 固定宽度
            double scale = (double)contentWidth.ValueInEmus / size.Width;

            //宽度缩放
            int width = (int)(size.Width * scale);
            int height = (int)(size.Height * scale);

            Run run = new Run();
            RunProperties runProperties = new RunProperties();

            Drawing drawing1 = new Drawing();

            wp.Inline inline1 = new wp.Inline() { 
                DistanceFromTop = (UInt32Value)0U, 
                DistanceFromBottom = (UInt32Value)0U, 
                DistanceFromLeft = (UInt32Value)0U, 
                DistanceFromRight = (UInt32Value)0U, 
                //AnchorId = "710459B9", EditId = "365C82D5" 
            };
            wp.Extent extent1 = new wp.Extent() { Cx = width, Cy = height };
            wp.EffectExtent effectExtent1 = new wp.EffectExtent()
            { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            wp.DocProperties docProperties1 = new wp.DocProperties() { Id = (UInt32Value)1U, Name = "Logo" };

            //wp.Anchor anchor1 = new wp.Anchor()
            //{
            //    DistanceFromTop = 0U,
            //    DistanceFromBottom = 0U,
            //    DistanceFromLeft = 114300U,
            //    DistanceFromRight = 114300U,
            //    SimplePos = false,
            //    RelativeHeight = 0U,
            //    BehindDoc = false,
            //    Locked = false,
            //    LayoutInCell = true,
            //    AllowOverlap = true
            //};

            //wp.SimplePosition simplePosition1 = new wp.SimplePosition()
            //{ X = 0, Y = 0 };

            //wp.HorizontalPosition horizontalPosition1 = new wp.HorizontalPosition() { RelativeFrom = wp.HorizontalRelativePositionValues.Column };
            //wp.PositionOffset positionOffset1 = new wp.PositionOffset() { Text = "0" };
            //horizontalPosition1.Append(positionOffset1);

            //wp.VerticalPosition verticalPosition1 = new wp.VerticalPosition() { RelativeFrom = wp.VerticalRelativePositionValues.Paragraph };
            //wp.PositionOffset positionOffset2 = new wp.PositionOffset() { Text = "0" };
            //verticalPosition1.Append(positionOffset2);

            wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 =
                new wp.NonVisualGraphicFrameDrawingProperties(new a.GraphicFrameLocks()
                {
                    NoChangeAspect = true,
                });

            string relationshipId = headerPart.GetIdOfPart(imagePart);
            a.Graphic graphic = new a.Graphic(
                new a.GraphicData(
                    new pic.Picture(
                        new pic.NonVisualPictureProperties
                        {
                            NonVisualDrawingProperties = new pic.NonVisualDrawingProperties()
                            { Id = (UInt32)headImgId++, Name = "Logo" },
                            NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                        },
                        new pic.BlipFill(
                            new a.Blip(
                                new a.ExtensionList(
                                    new a.BlipExtension(
                                        new a14.UseLocalDpi() { Val = false })
                                    {
                                        Uri = "{" + Guid.NewGuid() + "}"
                                    }
                                ))
                            { Embed = relationshipId },
                            new a.Stretch(
                                new a.FillRectangle())),
                        new pic.ShapeProperties(
                            new a.Transform2D(
                                new a.Offset() { X = 0L, Y = 0L },
                                new a.Extents() { Cx = width, Cy = height }),
                            new a.PresetGeometry(
                                new a.AdjustValueList()
                            )
                            { Preset = a.ShapeTypeValues.Rectangle }
                        )
                        { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

            //wp14.RelativeWidth relativeWidth1 = new wp14.RelativeWidth
            //{ ObjectId = wp14.SizeRelativeHorizontallyValues.Margin };
            //relativeWidth1.Append(new wp14.PercentageWidth
            //{
            //    Text = "0"
            //});

            //wp14.RelativeHeight relativeHeight1 = new wp14.RelativeHeight
            //{ RelativeFrom = wp14.SizeRelativeVerticallyValues.Margin };
            //relativeHeight1.Append(new wp14.PercentageHeight
            //{
            //    Text = "0"
            //});

            //anchor1.Append(simplePosition1);
            //anchor1.Append(horizontalPosition1);
            //anchor1.Append(verticalPosition1);
            //anchor1.Append(wrapSquare);
            //anchor1.Append(relativeWidth1);
            //anchor1.Append(relativeHeight1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic);

            drawing1.Append(inline1);

            run.Append(runProperties);
            run.Append(drawing1);
            return run;
        }

        /// <summary>
        /// 二维码
        /// </summary>
        /// <returns></returns>
        private Paragraph QRCodeParagraph(HeaderPart headerPart)
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "even" };
            Justification justification = new Justification() { Val = JustificationValues.Left };

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(justification);

            Paragraph paragraph = new Paragraph();

            Size size = new Size();

            ImagePart imagePart;
            if (DataUri.TryCreate(settings.QRCode, out DataUri dataUri))
            {
                ImagePrefetcher.knownContentType.TryGetValue(dataUri.Mime, out ImagePartType type);
                imagePart = headerPart.AddImagePart(type);
                imagePart.FeedData(new MemoryStream(dataUri.Data));
                using (var outputStream = imagePart.GetStream(FileMode.Create))
                {
                    outputStream.Write(dataUri.Data, 0, dataUri.Data.Length);
                    outputStream.Seek(0L, SeekOrigin.Begin);
                    size = ImagePrefetcher.GetImageSize(outputStream);
                }
            }
            else
            {
                return null;
            }

            double scale = (double)qrCodeSize.ValueInEmus / size.Height;
            //宽度缩放
            int width = (int)(size.Width * scale);
            int height = (int)(size.Height * scale);

            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            NoProof noProof = new NoProof();

            runProperties.Append(noProof);

            Drawing drawing1 = new Drawing();

            wp.Anchor anchor1 = new wp.Anchor()
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };

            wp.SimplePosition simplePosition1 = new wp.SimplePosition()
            { X = 0, Y = 0 };

            wp.HorizontalPosition horizontalPosition1 = new wp.HorizontalPosition
            { RelativeFrom = wp.HorizontalRelativePositionValues.Margin };
            horizontalPosition1.Append(new wp.HorizontalAlignment("right"));

            wp.VerticalPosition verticalPosition1 = new wp.VerticalPosition
            { RelativeFrom = wp.VerticalRelativePositionValues.TopMargin };
            verticalPosition1.Append(new wp.PositionOffset
            {
                Text = Convert.ToString(marginTop.ValueInEmus - qrCodeSize.ValueInEmus)
            });

            wp.Extent extent1 = new wp.Extent() { Cx = width, Cy = height };
            wp.EffectExtent effectExtent1 = new wp.EffectExtent()
            { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            wp.WrapNone wrapNone1 = new wp.WrapNone();
            wp.DocProperties docProperties1 = new wp.DocProperties() { Id = 1U, Name = "QRCode" };
            wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 =
                new wp.NonVisualGraphicFrameDrawingProperties(new a.GraphicFrameLocks()
                {
                    NoChangeAspect = true,
                });

            string relationshipId = headerPart.GetIdOfPart(imagePart);
            a.Graphic graphic = new a.Graphic(
                new a.GraphicData(
                    new pic.Picture(
                        new pic.NonVisualPictureProperties
                        {
                            NonVisualDrawingProperties = new pic.NonVisualDrawingProperties()
                            { Id = (UInt32)headImgId++, Name = "QRCode" },
                            NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                        },
                        new pic.BlipFill(
                            new a.Blip(
                                new a.ExtensionList(
                                    new a.BlipExtension(
                                        new a14.UseLocalDpi() { Val = false })
                                    {
                                        Uri = "{" + Guid.NewGuid() + "}"
                                    }
                                ))
                            { Embed = relationshipId },
                            new a.Stretch(
                                new a.FillRectangle())),
                        new pic.ShapeProperties(
                            new a.Transform2D(
                                new a.Offset() { X = 0L, Y = 0L },
                                new a.Extents() { Cx = width, Cy = height }),
                            new a.PresetGeometry(
                                new a.AdjustValueList()
                            )
                            { Preset = a.ShapeTypeValues.Rectangle }
                        )
                        { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

            wp14.RelativeWidth relativeWidth1 = new wp14.RelativeWidth
            { ObjectId = wp14.SizeRelativeHorizontallyValues.Margin };
            relativeWidth1.Append(new wp14.PercentageWidth
            {
                Text = "0"
            });

            wp14.RelativeHeight relativeHeight1 = new wp14.RelativeHeight
            { RelativeFrom = wp14.SizeRelativeVerticallyValues.Margin };
            relativeHeight1.Append(new wp14.PercentageHeight
            {
                Text = "0"
            });

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            run.Append(runProperties);
            run.Append(drawing1);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            return paragraph;
        }


        /// <summary>
        /// 水印
        /// </summary>
        /// <returns></returns>
        private Run WaterMark(HeaderPart headerPart)
        {
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Png);

            Size size;
            string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? string.Empty,
                "Resources",
                (isLandscape ? "landscape" : "portrait") + "_watermark.png");
            using (FileStream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                size = ImagePrefetcher.GetImageSize(stream);
                stream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(stream);
            }

            double scale = (double)contentHeight.ValueInEmus / size.Height;
            //宽度缩放
            int width = (int)(size.Width * scale);
            int height = (int)(size.Height * scale);

            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            //NoProof noProof = new NoProof();

            //runProperties.Append(noProof);

            Drawing drawing1 = new Drawing();

            wp.Anchor anchor1 = new wp.Anchor()
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };
            wp.SimplePosition simplePosition1 = new wp.SimplePosition() { X = 0L, Y = 0L };

            wp.HorizontalPosition horizontalPosition1 = new wp.HorizontalPosition() { RelativeFrom = wp.HorizontalRelativePositionValues.Margin };
            wp.HorizontalAlignment horizontalAlignment1 = new wp.HorizontalAlignment();
            horizontalAlignment1.Text = "center";
            horizontalPosition1.Append(horizontalAlignment1);

            wp.VerticalPosition verticalPosition1 = new wp.VerticalPosition() { RelativeFrom = wp.VerticalRelativePositionValues.Margin };
            wp.PositionOffset positionOffset1 = new wp.PositionOffset();
            positionOffset1.Text = "0";
            verticalPosition1.Append(positionOffset1);

            wp.Extent extent1 = new wp.Extent() { Cx = width, Cy = height };
            wp.EffectExtent effectExtent1 = new wp.EffectExtent()
            { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            wp.WrapNone wrapNone1 = new wp.WrapNone();
            wp.DocProperties docProperties1 = new wp.DocProperties() { Id = 1U, Name = "watermark" };
            wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 =
                new wp.NonVisualGraphicFrameDrawingProperties(new a.GraphicFrameLocks()
                {
                    NoChangeAspect = true,
                });

            string relationshipId = headerPart.GetIdOfPart(imagePart);
            a.Graphic graphic = new a.Graphic(
                new a.GraphicData(
                    new pic.Picture(
                        new pic.NonVisualPictureProperties
                        {
                            NonVisualDrawingProperties = new pic.NonVisualDrawingProperties()
                            { Id = (UInt32)headImgId++, Name = "watermark" },
                            NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                        },
                        new pic.BlipFill(
                            new a.Blip(
                                new a.ExtensionList(
                                    new a.BlipExtension(
                                        new a14.UseLocalDpi() { Val = false })
                                    {
                                        Uri = Guid.NewGuid().ToString(),
                                    }
                                ))
                            { Embed = relationshipId },
                            new a.Stretch(
                                new a.FillRectangle())),
                        new pic.ShapeProperties(
                            new a.Transform2D(
                                new a.Offset() { X = 0L, Y = 0L },
                                new a.Extents() { Cx = width, Cy = height }),
                            new a.PresetGeometry(
                                new a.AdjustValueList()
                            )
                            { Preset = a.ShapeTypeValues.Rectangle }
                        )
                        { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

            wp14.RelativeWidth relativeWidth1 = new wp14.RelativeWidth()
            { ObjectId = wp14.SizeRelativeHorizontallyValues.Margin };
            wp14.PercentageWidth percentageWidth1 = new wp14.PercentageWidth();
            percentageWidth1.Text = "0";
            relativeWidth1.Append(percentageWidth1);

            wp14.RelativeHeight relativeHeight1 = new wp14.RelativeHeight()
            { RelativeFrom = wp14.SizeRelativeVerticallyValues.Margin };
            wp14.PercentageHeight percentageHeight1 = new wp14.PercentageHeight();
            percentageHeight1.Text = "0";
            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            run.Append(runProperties);
            run.Append(drawing1);
            return run;
        }

        /// <summary>
        /// 装订线
        /// </summary>
        /// <returns></returns>
        private void Gutter()
        {
            // Paragraph firstParagraph = mainPart.Document.Body.GetFirstChild<Paragraph>();
            // ParagraphProperties pPr = firstParagraph.ParagraphProperties;
            // if (pPr == null)
            // {
            //     pPr = new ParagraphProperties();
            // }
            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Jpeg);
            Header header = NewHeader();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "default" };
            Justification justification = new Justification() { Val = JustificationValues.Left };

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(justification);

            Paragraph paragraph = new Paragraph();

            //左上方定位点
            double pointTop = settings.MarginTop - pointHeight - 0.2;
            pointLeft = settings.MarginLeft - pointWidth - 0.2;
            AnchorPoint leftTopPoint = new AnchorPoint(
                docPrId++,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                leftTopPoint.AddAnchorPoint()));

            //左下方定位点
            pointTop = settings.PageHeight - settings.MarginBottom + 0.2;
            pointLeft = settings.MarginLeft - pointWidth - 0.2;
            AnchorPoint leftBottomPoint = new AnchorPoint(
                docPrId++,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                leftBottomPoint.AddAnchorPoint()));
            //右上方定位点
            pointTop = settings.MarginTop - pointHeight - 0.2;
            pointLeft = settings.PageWidth - settings.MarginRight + 0.2;
            AnchorPoint rightTopPoint = new AnchorPoint(
                docPrId,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                rightTopPoint.AddAnchorPoint()));

            //右下方定位点
            pointTop = settings.PageHeight - settings.MarginBottom + 0.2;
            pointLeft = settings.PageWidth - settings.MarginRight + 0.2;
            AnchorPoint rightBottomPoint = new AnchorPoint(
                docPrId,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                rightBottomPoint.AddAnchorPoint()));
            // SectionProperties sPr = SectionProperties();

            Size size;
            string filepath = Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory ?? string.Empty,
                "Resources",
                "header_left.jpg");
            using (FileStream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                size = ImagePrefetcher.GetImageSize(stream);
                stream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(stream);
            }

            double scale = (double)pageHeight.ValueInEmus / size.Height;
            //宽度缩放
            int width = (int)(size.Width * scale);
            int height = (int)(size.Height * scale);

            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            NoProof noProof = new NoProof();

            runProperties.Append(noProof);

            Drawing drawing1 = new Drawing();

            wp.Anchor anchor1 = new wp.Anchor()
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };
            wp.SimplePosition simplePosition1 = new wp.SimplePosition() { X = 0L, Y = 0L };

            wp.HorizontalPosition horizontalPosition1 = new wp.HorizontalPosition()
            { RelativeFrom = wp.HorizontalRelativePositionValues.LeftMargin };
            wp.PositionOffset positionOffset1 = new wp.PositionOffset();
            positionOffset1.Text = "0";

            horizontalPosition1.Append(positionOffset1);

            wp.VerticalPosition verticalPosition1 = new wp.VerticalPosition()
            { RelativeFrom = wp.VerticalRelativePositionValues.TopMargin };
            wp.PositionOffset positionOffset2 = new wp.PositionOffset();
            positionOffset2.Text = "0";

            verticalPosition1.Append(positionOffset2);
            wp.Extent extent1 = new wp.Extent() { Cx = width, Cy = height };
            wp.EffectExtent effectExtent1 = new wp.EffectExtent()
            { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            wp.WrapNone wrapNone1 = new wp.WrapNone();
            wp.DocProperties docProperties1 = new wp.DocProperties() { Id = 1U, Name = "Gutter" };
            wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 =
                new wp.NonVisualGraphicFrameDrawingProperties(new a.GraphicFrameLocks()
                {
                    NoChangeAspect = true,
                });

            string relationshipId = headerPart.GetIdOfPart(imagePart);
            a.Graphic graphic = new a.Graphic(
                new a.GraphicData(
                    new pic.Picture(
                        new pic.NonVisualPictureProperties
                        {
                            NonVisualDrawingProperties = new pic.NonVisualDrawingProperties()
                            { Id = (UInt32)headImgId++, Name = "Gutter" },
                            NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                        },
                        new pic.BlipFill(
                            new a.Blip(
                                new a.ExtensionList(
                                    new a.BlipExtension(
                                        new a14.UseLocalDpi() { Val = false })
                                    {
                                        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                    }
                                ))
                            { Embed = relationshipId },
                            new a.Stretch(
                                new a.FillRectangle())),
                        new pic.ShapeProperties(
                            new a.Transform2D(
                                new a.Offset() { X = 0L, Y = 0L },
                                new a.Extents() { Cx = width, Cy = height }),
                            new a.PresetGeometry(
                                new a.AdjustValueList()
                            )
                            { Preset = a.ShapeTypeValues.Rectangle }
                        )
                        { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

            wp14.RelativeWidth relativeWidth1 = new wp14.RelativeWidth()
            { ObjectId = wp14.SizeRelativeHorizontallyValues.Margin };
            wp14.PercentageWidth percentageWidth1 = new wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            wp14.RelativeHeight relativeHeight1 = new wp14.RelativeHeight()
            { RelativeFrom = wp14.SizeRelativeVerticallyValues.Margin };
            wp14.PercentageHeight percentageHeight1 = new wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            run.Append(runProperties);
            run.Append(drawing1);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            header.Append(paragraph);


            headerPart.Header = header;
            string headerPartId = mainPart.GetIdOfPart(headerPart);

            foreach (SectionProperties sPr in sections)
            {
                // sPr.RemoveAllChildren<HeaderReference>();
                sPr.PrependChild(new HeaderReference()
                {
                    Id = headerPartId,
                    Type = HeaderFooterValues.Default
                });
            }
        }

        /// <summary>
        /// 密封线
        /// </summary>
        /// <returns></returns>
        private void Sealing()
        {
            // Paragraph firstParagraph = mainPart.Document.Body.GetFirstChild<Paragraph>();
            // ParagraphProperties pPr = firstParagraph.ParagraphProperties;
            // if (pPr == null)
            // {
            //     pPr = new ParagraphProperties();
            // }

            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            ImagePart imagePart = headerPart.AddImagePart(ImagePartType.Jpeg);
            Header header = NewHeader();
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "even" };
            Justification justification = new Justification() { Val = JustificationValues.Left };

            paragraphProperties.Append(paragraphStyleId);
            paragraphProperties.Append(justification);

            Paragraph paragraph = new Paragraph();

            //左上方定位点
            double pointTop = settings.MarginTop - pointHeight - 0.2;
            pointLeft = settings.MarginRight - pointWidth - 0.2;
            AnchorPoint leftTopPoint = new AnchorPoint(
                docPrId++,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                leftTopPoint.AddAnchorPoint()));

            //左下方定位点
            pointTop = settings.PageHeight - settings.MarginBottom + 0.2;
            pointLeft = settings.MarginRight - pointWidth - 0.2;
            AnchorPoint leftBottomPoint = new AnchorPoint(
                docPrId++,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                leftBottomPoint.AddAnchorPoint()));
            //右上方定位点
            pointTop = settings.MarginTop - pointHeight - 0.2;
            pointLeft = settings.PageWidth - settings.MarginLeft + 0.2;
            AnchorPoint rightTopPoint = new AnchorPoint(
                docPrId,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                rightTopPoint.AddAnchorPoint()));

            //右下方定位点
            pointTop = settings.PageHeight - settings.MarginBottom + 0.2;
            pointLeft = settings.PageWidth - settings.MarginLeft + 0.2;
            AnchorPoint rightBottomPoint = new AnchorPoint(
                docPrId,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                rightBottomPoint.AddAnchorPoint()));

            Size size;
            string filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? string.Empty,
                "Resources",
                "header_right.jpg");
            using (FileStream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                size = ImagePrefetcher.GetImageSize(stream);
                stream.Seek(0, SeekOrigin.Begin);
                imagePart.FeedData(stream);
            }

            double scale = (double)pageHeight.ValueInEmus / size.Height;
            //宽度缩放
            int width = (int)(size.Width * scale);
            int height = (int)(size.Height * scale);

            Run run = new Run();

            RunProperties runProperties = new RunProperties();
            NoProof noProof = new NoProof();

            runProperties.Append(noProof);

            Drawing drawing1 = new Drawing();

            wp.Anchor anchor1 = new wp.Anchor()
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = false,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            };
            wp.SimplePosition simplePosition1 = new wp.SimplePosition() { X = 0L, Y = 0L };

            wp.HorizontalPosition horizontalPosition1 = new wp.HorizontalPosition()
            { RelativeFrom = wp.HorizontalRelativePositionValues.RightMargin };
            wp.HorizontalAlignment horizontalAlignment1 = new wp.HorizontalAlignment();
            horizontalAlignment1.Text = "right";
            horizontalPosition1.Append(horizontalAlignment1);

            wp.VerticalPosition verticalPosition1 = new wp.VerticalPosition()
            { RelativeFrom = wp.VerticalRelativePositionValues.TopMargin };
            wp.PositionOffset positionOffset2 = new wp.PositionOffset();
            positionOffset2.Text = "0";

            verticalPosition1.Append(positionOffset2);
            wp.Extent extent1 = new wp.Extent() { Cx = width, Cy = height };
            wp.EffectExtent effectExtent1 = new wp.EffectExtent()
            { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            wp.WrapNone wrapNone1 = new wp.WrapNone();
            wp.DocProperties docProperties1 = new wp.DocProperties() { Id = 1U, Name = "Gutter" };
            wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 =
                new wp.NonVisualGraphicFrameDrawingProperties(new a.GraphicFrameLocks()
                {
                    NoChangeAspect = true,
                });


            string relationshipId = headerPart.GetIdOfPart(imagePart);
            a.Graphic graphic = new a.Graphic(
                new a.GraphicData(
                    new pic.Picture(
                        new pic.NonVisualPictureProperties
                        {
                            NonVisualDrawingProperties = new pic.NonVisualDrawingProperties()
                            { Id = (UInt32)headImgId++, Name = "Gutter" },
                            NonVisualPictureDrawingProperties = new pic.NonVisualPictureDrawingProperties(
                                new a.PictureLocks() { NoChangeAspect = true, NoChangeArrowheads = true })
                        },
                        new pic.BlipFill(
                            new a.Blip(
                                new a.ExtensionList(
                                    new a.BlipExtension(
                                        new a14.UseLocalDpi() { Val = false })
                                    {
                                        Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                    }
                                ))
                            { Embed = relationshipId },
                            new a.Stretch(
                                new a.FillRectangle())),
                        new pic.ShapeProperties(
                            new a.Transform2D(
                                new a.Offset() { X = 0L, Y = 0L },
                                new a.Extents() { Cx = width, Cy = height }),
                            new a.PresetGeometry(
                                new a.AdjustValueList()
                            )
                            { Preset = a.ShapeTypeValues.Rectangle }
                        )
                        { BlackWhiteMode = a.BlackWhiteModeValues.Auto })
                )
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });

            wp14.RelativeWidth relativeWidth1 = new wp14.RelativeWidth()
            { ObjectId = wp14.SizeRelativeHorizontallyValues.Margin };
            wp14.PercentageWidth percentageWidth1 = new wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            wp14.RelativeHeight relativeHeight1 = new wp14.RelativeHeight()
            { RelativeFrom = wp14.SizeRelativeVerticallyValues.Margin };
            wp14.PercentageHeight percentageHeight1 = new wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            run.Append(runProperties);
            run.Append(drawing1);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            header.Append(paragraph);


            headerPart.Header = header;
            string headerPartId = mainPart.GetIdOfPart(headerPart);

            foreach (SectionProperties sPr in sections)
            {
                // sPr.RemoveAllChildren<HeaderReference>();
                sPr.PrependChild(new HeaderReference()
                {
                    Id = headerPartId,
                    Type = HeaderFooterValues.Even
                });
            }

            // sPr.PrependChild(new HeaderReference()
            // {
            //     Id = headerPartId, Type = HeaderFooterValues.Default
            // });
            //
            // pPr.Append(sPr);
            // firstParagraph.ParagraphProperties = pPr;
        }

        /// <summary>
        /// 生成页眉部件内容
        /// </summary>
        private void GenerateHeaderPartContent()
        {
            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            Header header = NewHeader();
            int headerCount = 0;

            ParagraphProperties paragraphProperties = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "default" };
            paragraphProperties.Append(paragraphStyleId);

            Paragraph paragraph = new Paragraph();

            paragraph.Append(paragraphProperties);

            if (settings.Logo)
            {
                Run run = Logo(headerPart);
                if (paragraph != null)
                {
                    paragraph.Append(run);
                    headerCount++;
                }
            }
            if (settings.WaterMark)
            {
                Run run = WaterMark(headerPart);
                if (paragraph != null)
                {
                    paragraph.Append(run);
                    headerCount++;
                }
            }

            if (headerCount == 0)
            {
                return;
            }
            header.Append(paragraph);
            headerPart.Header = header;
            string headerPartId = mainPart.GetIdOfPart(headerPart);

            foreach (SectionProperties sPr in sections)
            {
                // sPr.RemoveAllChildren<HeaderReference>();
                sPr.PrependChild(new HeaderReference()
                {
                    Id = headerPartId,
                    Type = HeaderFooterValues.Default
                });
            }
            //if (settings.Logo != String.Empty)
            //{
            //    Logo();
            //}else  if (settings.HeaderMode == 1)
            //{
            //    ThreeColumnsHeader();
            //}
            //else if (settings.HasHeader == 1)
            //{
            //    pointWidthUnit = new Unit(UnitMetric.Centimeter, pointWidth);
            //    pointHeightUnit = new Unit(UnitMetric.Centimeter, pointHeight);
            //    //设置密封线
            //    if (settings.SealLine)
            //    {
            //        Gutter();
            //        Sealing();

            //        //设置页眉奇偶页不同
            //        docSettings.Append(new EvenAndOddHeaders());
            //    }
            //    else if (settings.HasAnchorPoint)
            //    {
            //        PositionPoint();
            //    }
            //}
            //else if (settings.WaterMark)
            //{
            //    WaterMark(null);
            //}
        }

        /// <summary>
        /// 定位点
        /// </summary>
        private void PositionPoint()
        {
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            Header header = NewHeader();
            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

            Paragraph paragraph = new Paragraph();

            ParagraphProperties pPr = new ParagraphProperties();
            ParagraphStyleId pStyle = new ParagraphStyleId() { Val = "Header" };

            pPr.Append(pStyle);
            paragraph.Append(pPr);

            //左上方定位点
            double pointTop = settings.MarginTop - pointHeight - 0.2;
            pointLeft = settings.MarginLeft - pointWidth - 0.2;
            AnchorPoint leftTopPoint = new AnchorPoint(
                docPrId++,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                leftTopPoint.AddAnchorPoint()));

            //左下方定位点
            pointTop = settings.PageHeight - settings.MarginBottom + 0.2;
            pointLeft = settings.MarginLeft - pointWidth - 0.2;
            AnchorPoint leftBottomPoint = new AnchorPoint(
                docPrId++,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                leftBottomPoint.AddAnchorPoint()));

            //右上方定位点
            pointTop = settings.MarginTop - pointHeight - 0.2;
            pointLeft = settings.PageWidth - settings.MarginRight + 0.2;
            AnchorPoint rightTopPoint = new AnchorPoint(
                docPrId,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                rightTopPoint.AddAnchorPoint()));

            //右下方定位点
            pointTop = settings.PageHeight - settings.MarginBottom + 0.2;
            pointLeft = settings.PageWidth - settings.MarginRight + 0.2;
            AnchorPoint rightBottomPoint = new AnchorPoint(
                docPrId,
                pointWidthUnit,
                pointHeightUnit,
                new Unit(UnitMetric.Centimeter, pointLeft),
                new Unit(UnitMetric.Centimeter, pointTop));
            paragraph.Append(new Run(
                new RunProperties(
                    new NoProof()),
                rightBottomPoint.AddAnchorPoint()));

            header.Append(paragraph);

            headerPart.Header = header;
            string headerPartId = mainPart.GetIdOfPart(headerPart);

            foreach (SectionProperties sPr in sections)
            {
                sPr.PrependChild(new HeaderReference()
                {
                    Id = headerPartId,
                    Type = HeaderFooterValues.Default
                });
            }
        }


        /// <summary>
        /// 二维码
        /// </summary>
        /// <returns></returns>
        private void ThreeColumnsHeader()
        {
            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();
            HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();
            Header header = NewHeader();
            ParagraphProperties paragraphProperties = new ParagraphProperties()
            {
                ParagraphStyleId = new ParagraphStyleId() { Val = "even" }
            };

            Paragraph paragraph = new Paragraph()
            {
                ParagraphProperties = paragraphProperties
            };

            RunProperties runProperties = new RunProperties
            {
                RunFonts = new RunFonts
                {
                    Ascii = "Times New Roman",
                    HighAnsi = "Times New Roman",
                    EastAsia = "宋体"
                },
                FontSize = new FontSize
                {
                    Val = "21"
                },
                FontSizeComplexScript = new FontSizeComplexScript
                {
                    Val = "21"
                }
            };
            // 左边文字
            Run leftRun = new Run
            {
                RunProperties = runProperties
            };
            leftRun.Append(new Text(settings.ThreeColumnsHeader.Left));
            paragraph.Append(leftRun);

            // 左边Tab
            PositionalTab leftTab = new PositionalTab
            {
                Alignment = AbsolutePositionTabAlignmentValues.Center,
                RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin,
                Leader = AbsolutePositionTabLeaderCharValues.None
            };
            paragraph.Append(leftTab);
            // 中间文字
            Run centerRun = new Run()
            {
                RunProperties = (RunProperties)runProperties.Clone()
            };
            centerRun.Append(new Text(settings.ThreeColumnsHeader.Center));
            paragraph.Append(centerRun);

            // 右边Tab
            PositionalTab rightTab = new PositionalTab
            {
                Alignment = AbsolutePositionTabAlignmentValues.Right,
                RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin,
                Leader = AbsolutePositionTabLeaderCharValues.None
            };
            paragraph.Append(rightTab);

            // 右边文字
            Run rightRun = new Run()
            {
                RunProperties = (RunProperties)runProperties.Clone()
            };
            rightRun.Append(new Text(settings.ThreeColumnsHeader.Right));
            paragraph.Append(rightRun);

            header.Append(paragraph);

            headerPart.Header = header;
            string headerPartId = mainPart.GetIdOfPart(headerPart);

            foreach (SectionProperties sPr in sections)
            {
                // sPr.RemoveAllChildren<HeaderReference>();
                sPr.PrependChild(new HeaderReference()
                {
                    Id = headerPartId,
                    Type = HeaderFooterValues.Default
                });
            }
        }

        /// <summary>
        /// 生成页脚部件内容
        /// </summary>
        /// <param name="part">页脚部件</param>
        private void GenerateFooterPartContent(FooterPart part)
        {
            Footer footer1 = new Footer()
            { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            footer1.AddNamespaceDeclaration("wpc",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp14",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("wp",
                "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer1.AddNamespaceDeclaration("wpg",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer1.AddNamespaceDeclaration("wpi",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer1.AddNamespaceDeclaration("wps",
                "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph = new Paragraph()
            { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties1.Append(paragraphStyleId1);
            paragraphProperties1.Justification = new Justification
            {
                Val = JustificationValues.Center
            };
            paragraph.Append(paragraphProperties1);
            if (!settings.IgnorePageNumber)
            {
                RunProperties rPr = new RunProperties()
                {
                    RunFonts = new RunFonts
                    {
                        Ascii = "仿宋",
                        HighAnsi = "仿宋",
                        EastAsia = "仿宋"
                    },
                    FontSize = new FontSize
                    {
                        Val = "18"
                    },
                    FontSizeComplexScript = new FontSizeComplexScript
                    {
                        Val = "18"
                    }
                };
                Run runs = new Run(rPr);
                runs.Append(new Text("刷题、找试卷就用考霸疯狂刷题APP（第"));
                runs.Append(new FieldChar
                {
                    FieldCharType = FieldCharValues.Begin,
                });
                runs.Append(new FieldCode(@" PAGE \* MERGEFORMAT ")
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
                runs.Append(new FieldChar
                {
                    FieldCharType = FieldCharValues.Separate
                });
                runs.Append(new FieldChar
                {
                    FieldCharType = FieldCharValues.End
                });
                runs.Append(new Text("页/共"));
                runs.Append(new FieldChar
                {
                    FieldCharType = FieldCharValues.Begin,
                });
                runs.Append(new FieldCode(@" NUMPAGES \* MERGEFORMAT ")
                {
                    Space = SpaceProcessingModeValues.Preserve
                });
                runs.Append(new FieldChar
                {
                    FieldCharType = FieldCharValues.Separate
                });
                runs.Append(new FieldChar
                {
                    FieldCharType = FieldCharValues.End
                });
                runs.Append(new Text("页）"));
                paragraph.Append(runs);
            }

            footer1.Append(paragraph);

            part.Footer = footer1;
        }

        public void Setting()
        {
            DocumentSettingsPart settingsPart = mainPart.DocumentSettingsPart;
            if (settingsPart == null)
            {
                settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            }

            docSettings = settingsPart.Settings;
            if (docSettings == null)
            {
                docSettings = NewSettings();
            }

            Compatibility compat = new Compatibility(
                new SpaceForUnderline(),
                new BalanceSingleByteDoubleByteWidth(),
                new DoNotLeaveBackslashAlone(),
                new UnderlineTrailingSpaces(),
                new DoNotExpandShiftReturn(),
                new AdjustLineHeightInTable(),
                new UseFarEastLayout(),
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.CompatibilityMode,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "15",
                },
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.OverrideTableStyleFontSizeAndJustification,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "1",
                },
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.EnableOpenTypeFeatures,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "1",
                },
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.DoNotFlipMirrorIndents,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "1",
                },
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.DifferentiateMultirowTableHeaders,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "1",
                },
                new CompatibilitySetting()
                {
                    Name = CompatSettingNameValues.UseWord2013TrackBottomHyphenation,
                    Uri = "http://schemas.microsoft.com/office/word",
                    Val = "0",
                }
            );
            docSettings.Append(compat);

            PageSetup();
            HeaderFooter();
            settingsPart.Settings = docSettings;
        }
    }
}