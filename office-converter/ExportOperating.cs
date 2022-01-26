using Library.Model;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.Word.Shape;

namespace OfficeConverter
{
    /// <summary>
    /// 导出操作
    /// </summary>
    public class ExportOperating
    {
        private readonly ExportWordSetting setting;
        private Application word;
        private Document doc;
        private Global wordBasic;
        private List<float> columnWidths = new List<float>();
        private string dstFile;

        public ExportOperating(string srcFile, string dstFile, ExportWordSetting setting)
        {
            wordBasic = new Global();
            this.setting = setting;
            this.dstFile = dstFile;

            string[] w = string.IsNullOrEmpty(this.setting.ColumnWidth)
                ? new string[0]
                : this.setting.ColumnWidth.Split(',');
            foreach (string width in w)
            {
                columnWidths.Add(wordBasic.CentimetersToPoints((float) (Convert.ToDouble(width))));
            }

            word = new Application();
            Documents docs = word.Documents;
            doc = docs.Open(srcFile);
            doc.Activate();
        }

        /// <summary>
        /// 分栏处理
        /// </summary>
        public void Columns()
        {
            if (setting.EvenlySpaced == 1 || setting.Columns == 1)
            {
                //栏平均
                doc.PageSetup.TextColumns.SetCount(setting.Columns);

                if (setting.Columns > 1)
                {
                    doc.PageSetup.TextColumns.Spacing = setting.ColumnSpacing;
                    doc.PageSetup.TextColumns.EvenlySpaced = -1;
                    doc.PageSetup.TextColumns.LineBetween = setting.HasLine;
                }
            }
            else
            {
                //多栏的处理：必须指定其他列的宽度
                float width = doc.PageSetup.PageWidth;
                float widthRight = 0f;
                foreach (float f in columnWidths)
                {
                    widthRight += f;
                }

                float widthLeft = width - ((setting.Columns - 1) * setting.ColumnSpacing) - widthRight;
                float widthAvg = (setting.Columns == (columnWidths.Count - 1))
                    ? 0
                    : (widthLeft / (setting.Columns - columnWidths.Count));
                int currentCount = columnWidths.Count;
                for (int i = currentCount; i < setting.Columns - 1; i++)
                {
                    columnWidths.Add(widthAvg);
                }

                doc.PageSetup.TextColumns.SetCount(1);
                object columnWidth = 0f;
                object thisSpacing = setting.ColumnSpacing;
                object eve = false;
                for (int i = 0; i < setting.Columns - 1; i++)
                {
                    columnWidth = columnWidths[i];
                    thisSpacing = setting.ColumnSpacing;
                    eve = false;
                    doc.PageSetup.TextColumns.Add(ref columnWidth, ref thisSpacing, ref eve);
                }

                doc.PageSetup.TextColumns.LineBetween = setting.HasLine;
            }
        }

        /// <summary>
        /// 页码，页眉，页脚
        /// </summary>
        public void HeaderFooter()
        {
            object nothing = Missing.Value;
            object saveWithDocument = true;
            int pages = doc.ComputeStatistics(WdStatistic.wdStatisticPages, ref nothing);
            if (setting.HasHeader == 1)
            {
                //装订线的处理：奇偶页不同，左右对称
                doc.PageSetup.OddAndEvenPagesHeaderFooter = -1; //奇偶页不同
                word.ActiveWindow.View.Type = WdViewType.wdPrintView;

                word.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryHeader; //进入页眉设置，奇数页，其中页眉边距在页面设置中已完成
                word.ActiveWindow.Selection.HeaderFooter.LinkToPrevious = false;
                //插入页眉图片 
                word.ActiveWindow.Selection.ParagraphFormat.Alignment =
                    WdParagraphAlignment.wdAlignParagraphLeft; //页眉中的文字左对齐
                string headerFile =
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory ?? "./", "Resources", "OddHeaderFile.jpg");
                if (!File.Exists(headerFile))
                {
                    Console.WriteLine("文件不存在，请确保安装正确！");
                    return;
                }

                InlineShape shape1 =
                    word.ActiveWindow.ActivePane.Selection.InlineShapes.AddPicture(headerFile, ref nothing,
                        ref saveWithDocument, ref nothing);
                Shape shape = shape1.ConvertToShape();

                float shapeHeight = shape.Height;
                float pageHeight = doc.PageSetup.PageHeight;

                float scale = pageHeight / shapeHeight;
                //宽度缩放
                shape.ScaleWidth(scale, MsoTriState.msoFalse);
                //高度缩放
                shape.ScaleHeight(scale, MsoTriState.msoFalse);
                //相对于左边距 https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.word.wdrelativehorizontalposition?view=word-pia
                shape.RelativeHorizontalPosition =
                    WdRelativeHorizontalPosition.wdRelativeHorizontalPositionLeftMarginArea;
                //相对于上边距 https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.word.wdrelativeverticalposition?view=word-pia
                shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionTopMarginArea;
                shape.Left = 0; // new Microsoft.Office.Interop.Word.Global().CentimetersToPoints(38.2f);
                shape.Top = 0;
                if (pages > 1)
                {
                    word.ActiveWindow.View.SeekView = WdSeekView.wdSeekEvenPagesHeader; //进入页眉设置，偶数页，其中页眉边距在页面设置中已完成
                    word.ActiveWindow.Selection.HeaderFooter.LinkToPrevious = false;
                    word.ActiveWindow.Selection.ParagraphFormat.Alignment =
                        WdParagraphAlignment.wdAlignParagraphRight; //页眉中的文字右对齐

                    //插入页眉图片
                    headerFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "EvenHeadFile.jpg");

                    if (!File.Exists(headerFile))
                    {
                        Console.WriteLine("文件不存在，请确保安装正确！");
                        return;
                    }

                    shape1 = word.ActiveWindow.ActivePane.Selection.InlineShapes.AddPicture(headerFile, ref nothing,
                        ref saveWithDocument, ref nothing);
                    shape = shape1.ConvertToShape();
                    shapeHeight = shape.Height;
                    scale = pageHeight / shapeHeight;
                    //宽度缩放
                    shape.ScaleWidth(scale, MsoTriState.msoFalse);
                    //高度缩放
                    shape.ScaleHeight(scale, MsoTriState.msoFalse);
                    //相对于上边距 https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.word.wdrelativeverticalposition?view=word-pia
                    shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionTopMarginArea;
                    //相对于左边距 https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.word.wdrelativehorizontalposition?view=word-pia
                    shape.RelativeHorizontalPosition =
                        WdRelativeHorizontalPosition.wdRelativeHorizontalPositionLeftMarginArea;
                    shape.Left =
                        doc.PageSetup.PageWidth -
                        shape.Width; // new Microsoft.Office.Interop.Word.Global().CentimetersToPoints(38.2f);
                    shape.Top = 0;
                }

                //去掉页眉的横线
                word.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle =
                    WdLineStyle.wdLineStyleNone;
                word.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
            }
            else if (setting.HasHeader == 2)
            {
                //订正线的处理：找到答案部分，分为2节。第一节添加装订线。
                //????
                doc.PageSetup.OddAndEvenPagesHeaderFooter = 0; //奇偶页相同
                //找特殊文字
                word.ActiveWindow.Selection.Find.ClearFormatting();

                word.ActiveWindow.Selection.Find.Text = "【答案与解析】";
                word.ActiveWindow.Selection.Find.Forward = true;
                word.ActiveWindow.Selection.Find.MatchCase = false;
                if (word.ActiveWindow.Selection.Find.Execute())
                {
                    object unit = WdUnits.wdCharacter;
                    object count = 1;
                    word.ActiveWindow.Selection.MoveLeft(ref unit, ref count);
                    object breakType = WdBreakType.wdSectionBreakNextPage;
                    word.ActiveWindow.Selection.InsertBreak(ref breakType);
                }

                word.ActiveWindow.View.SeekView = WdSeekView.wdSeekCurrentPageHeader; //进入页眉设置，偶数页，其中页眉边距在页面设置中已完成
                word.ActiveWindow.Selection.HeaderFooter.LinkToPrevious = false;
                word.ActiveWindow.Selection.ParagraphFormat.Alignment =
                    WdParagraphAlignment.wdAlignParagraphRight; //页眉中的文字右对齐
                string headerFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources",
                    "CorrectLineFile.png");

                if (!File.Exists(headerFile))
                {
                    Console.WriteLine("文件不存在，请确保安装正确！");
                    return;
                }

                word.ActiveWindow.View.Type = WdViewType.wdPrintView;
                InlineShape shape1 =
                    word.ActiveWindow.ActivePane.Selection.InlineShapes.AddPicture(headerFile, ref nothing,
                        ref saveWithDocument, ref nothing);
                Shape shape = shape1.ConvertToShape();
                shape.Left =
                    -1 * doc.PageSetup
                        .LeftMargin; // new Microsoft.Office.Interop.Word.Global().CentimetersToPoints(38.2f);
                shape.Top = -1 * doc.PageSetup.TopMargin;
                shape.Width = doc.PageSetup.PageWidth;
                shape.Height = doc.PageSetup.PageHeight;
                //去掉页眉的横线
                word.ActiveWindow.ActivePane.Selection.ParagraphFormat.Borders[WdBorderType.wdBorderBottom].LineStyle =
                    WdLineStyle.wdLineStyleNone;
                word.ActiveWindow.ActivePane.Selection.Borders[WdBorderType.wdBorderBottom].Visible = false;
            }

            //页脚
            word.ActiveWindow.View.SeekView = WdSeekView.wdSeekPrimaryFooter; //进入页脚设置，奇数页
            word.ActiveWindow.Selection.HeaderFooter.LinkToPrevious = false;
            word.Selection.TypeText("第");
            object wdFieldEmpty = WdFieldType.wdFieldEmpty;
            object preserveFormatting = true;
            object fieldText = "PAGE";
            word.Selection.Fields.Add(word.Selection.Range, ref wdFieldEmpty, ref fieldText, ref preserveFormatting);
            word.Selection.TypeText("页/共");
            object numPages = "NUMPAGES";
            word.Selection.Fields.Add(word.Selection.Range, ref wdFieldEmpty, ref numPages, ref preserveFormatting);
            word.Selection.TypeText("页");
            word.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            if (pages > 1 && setting.HasHeader == 1)
            {
                word.ActiveWindow.View.SeekView = WdSeekView.wdSeekEvenPagesFooter; //进入页脚设置，偶数页
                word.ActiveWindow.Selection.HeaderFooter.LinkToPrevious = false;
                word.Selection.TypeText("第");
                word.Selection.Fields.Add(word.Selection.Range, ref wdFieldEmpty, ref fieldText,
                    ref preserveFormatting);
                word.Selection.TypeText("页/共");
                word.Selection.Fields.Add(word.Selection.Range, ref wdFieldEmpty, ref numPages,
                    ref preserveFormatting);
                word.Selection.TypeText("页");
                word.Selection.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            }

            word.ActiveWindow.ActivePane.View.SeekView = WdSeekView.wdSeekMainDocument; //退出页眉设置

            word.ActiveWindow.View.Type = WdViewType.wdNormalView;
        }

        public void Save()
        {
            doc.Save();
        }

        /// <summary>
        /// 保存并退出
        /// </summary>
        public void CloseAndQuit()
        {
            doc.Close();
            word.Quit();
        }

        public void ConvertPDF()
        {
            doc.SaveAs2(dstFile, WdSaveFormat.wdFormatPDF);
        }

        public void Export()
        {
            // PageSetup();
            // Console.WriteLine(@"PageSetup:{0}", DateTime.Now.ToString());
            // Columns();
            // Console.WriteLine(@"Columns:{0}", DateTime.Now.ToString());
            // HeaderFooter();
            // Console.WriteLine(@"HeaderFooter:{0}", DateTime.Now.ToString());
            // if (setting.Convert2PDF == "1")
            // {
            //
            // }
            // else
            // {
            //     Save();
            //     Console.WriteLine(@"Save:{0}", DateTime.Now.ToString());
            // }
            
            Console.WriteLine(@"OpenWord:{0}", DateTime.Now.ToString());
            ConvertPDF();
            Console.WriteLine(@"ConvertPDF:{0}", DateTime.Now.ToString());
            CloseAndQuit();
            Console.WriteLine(@"CloseAndQuit:{0}", DateTime.Now.ToString());
        }
    }
}