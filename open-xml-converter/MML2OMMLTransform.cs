using System;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Xsl;

namespace OpenXmlConverter
{
    /// <summary>
    /// 
    /// </summary>
    public class MML2OMMLTransform
    {
        // 定义一个静态变量来保存类的实例
        private static MML2OMMLTransform uniqueInstance;

        // 定义一个标识确保线程同步
        private static readonly object locker = new object();

        private XslCompiledTransform xslTransform;

        private XmlWriterSettings settings;

        // 定义私有构造函数，使外界不能创建该类实例
        private MML2OMMLTransform()
        {
            //https://stackoverflow.com/questions/10993621/openxml-sdk-and-mathml
            xslTransform = new XslCompiledTransform();

            // The MML2OMML.xsl file is located under 
            // %ProgramFiles%\Microsoft Office\Office12\
            string xslPath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "Resources", "MML2OMML.XSL");
            xslTransform.Load(xslPath);
            settings = xslTransform.OutputSettings.Clone();

            // Configure xml writer to omit xml declaration.
            settings.ConformanceLevel = ConformanceLevel.Fragment;
            settings.OmitXmlDeclaration = true;
            settings.CheckCharacters = false;
        }

        /// <summary>
        /// 定义公有方法提供一个全局访问点,同时你也可以定义公有属性来提供全局访问点
        /// </summary>
        /// <returns></returns>
        public static MML2OMMLTransform GetInstance()
        {
            // 当第一个线程运行到这里时，此时会对locker对象 "加锁"，
            // 当第二个线程运行该方法时，首先检测到locker对象为"加锁"状态，该线程就会挂起等待第一个线程解锁
            // lock语句运行完之后（即线程运行完之后）会对该对象"解锁"
            // 双重锁定只需要一句判断就可以了
            if (uniqueInstance == null)
            {
                lock (locker)
                {
                    // 如果类的实例不存在则创建，否则直接返回
                    if (uniqueInstance == null)
                    {
                        uniqueInstance = new MML2OMMLTransform();
                    }
                }
            }

            return uniqueInstance;
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <param name="mathml"></param>
        /// <returns></returns>
        public string Transform(string mathml)
        {
            string officeML;
            using (XmlReader reader =
                XmlReader.Create(new StringReader(mathml)))
            {
                using (MemoryStream ms = new MemoryStream())
                {
                    XmlWriter xw = XmlWriter.Create(ms, settings);

                    // Transform our MathML to OfficeMathML
                    xslTransform.Transform(reader, xw);
                    ms.Seek(0, SeekOrigin.Begin);

                    StreamReader sr = new StreamReader(ms, Encoding.UTF8);

                    officeML = sr.ReadToEnd();
                }
            }

            return officeML;
        }
    }
}