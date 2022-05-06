namespace Library.Model
{
    public class ExportWordSetting
    {
        public string Convert2PDF { get; set; } // 是否抓换为PDF 1表示转换为PDF

        // public int Orientation { get; set; } // 纸张方向 0竖向1横向
        // public int PaperSize { get; set; } // 纸张大小 -1代表自定义纸张
        public float MarginLeft { get; set; } // 左边距 单位是cm（直接传数字，不要+cm字符）
        public float MarginTop { get; set; } // 上边距 单位是cm（直接传数字，不要+cm字符）
        public float MarginRight { get; set; } // 右边距 单位是cm（直接传数字，不要+cm字符）
        public float MarginBottom { get; set; } // 下边距 单位是cm（直接传数字，不要+cm字符）
        public int HasHeader { get; set; } // 是否有页眉，0没有，1装订线，2订正线
        public double PageWidth { get; set; } // 纸张宽 单位是cm（直接传数字，不要+cm字符）
        public double PageHeight { get; set; } // 纸张高 单位是cm（直接传数字，不要+cm字符）
        public int MirrorMargin { get; set; } // 0表示不对称,非0表示对称
        public int Columns { get; set; } // 每页分栏数 本参数 >1且 EvenlySpaced=0 多栏且不平均分栏，其他分栏参数才有用。
        public int EvenlySpaced { get; set; } // 是否是平均分栏，1-是平均，0是不平均
        public int ColumnSpacing { get; set; } // 分栏的间隔大小，单位是字符
        public int HasLine { get; set; } // 分栏是否有分隔竖线，0无分隔线，非0有分隔线
        public bool HasAnchorPoint { get; set; } //是否包含定位点
        public bool SealLine { get; set; } // 密封线
        public bool WaterMark { get; set; } // 水印

        // 分栏宽度集合示例：
        // 分2栏，右栏5.5cm。 columns=2&columnWidth=5.5
        // 分3栏，中间栏5.5，右栏5.5。 columns=3&columnWidth=5.5,5.5
        // 分3栏，中间栏5.5，其他2栏平均分配。 columns=3&columnWidth=5.5
        public string ColumnWidth { get; set; } // 分栏的宽度，单位为cm。除了第一栏之外的宽度集合，用逗号分隔。

        public bool Logo { get; set; } //学校Logo

        public bool IgnorePageNumber { get; set; } //忽略页码

        public string QRCode { get; set; } //二维码
        public string GradeName { get; set; }//年级名
        public string SubjectName { get; set; }//学科名
        public string FooterContent { get; set; }//页脚内容
        public bool CancelSnapToGrid { get; set; }   //取消段落对齐到网格

        public int HeaderMode { get; set; } //页面显示模式
        // 1 三栏头显示 学号-姓名-班级    考试名称	   考试时间:21:10:12
        public ThreeColumnsHeader ThreeColumnsHeader { get; set; }

    }

    public class ThreeColumnsHeader
    {
        public string Left { get; set; }
        public string Center { get; set; }
        public string Right { get; set; }
    }

    public class HTML2WordParam
    {
        public string HTML { get; set; }
        public string Filepath { get; set; } // word文件要保存到的路径
        public string PDFPath { get; set; } // PDF文件要保存到的路径
        public ExportWordSetting Setting { get; set; }
    }

    public class HTML2WordResponse
    {
        public string WordPath { get; set; }
        public string PdfPath { get; set; }
    }
}