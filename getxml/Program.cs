using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Visio;
using System.IO;
using System.Xml;
using System.Collections;

namespace getxml
{
    class Program
    {
        static void Main(string[] args)
        {
            //初始化一个app，运行Visio程序
            ApplicationClass app = new ApplicationClass();

            //获取UML图路径 
            string path = System.IO.Directory.GetCurrentDirectory(); 
            String filePath = path + "\\uml\\test1.vsdx";

            //以只读形式打开一个Visio文件
            Document doc;
            doc = app.Documents.OpenEx(filePath, (short)Microsoft.Office.Interop.Visio.VisOpenSaveArgs.visOpenCopy);
            String outputDir = path + "\\output";
            Directory.GetDirectories(outputDir);

            try
            {
                XmlTextWriter textwriter = new XmlTextWriter(outputDir + "\\test1.xml", null);
                //Open the document
                textwriter.WriteStartDocument();
                //Write comments
                textwriter.WriteComment("Store xml information in this xml file");
                //Start Page
                textwriter.WriteStartElement("page");
                ArrayList list = new ArrayList();
                Page page = doc.Pages.get_ItemFromID(0);

                //Write page attribute width
                textwriter.WriteStartAttribute("width");
                textwriter.WriteValue((int)(page.PageSheet.get_Cells("PageWidth").ResultIU * 96));
                textwriter.WriteEndAttribute();

                //Write page attribute height
                textwriter.WriteStartAttribute("height");
                textwriter.WriteValue((int)(page.PageSheet.get_Cells("PageHeight").ResultIU * 96));
                textwriter.WriteEndAttribute();

                foreach (Shape shape in page.Shapes)
                {
                    String id = shape.Text;

                    //输出名称不为空的Shape对象
                    if (id.Length > 0)
                    {
                        ////Start shape
                        textwriter.WriteStartElement("shape");

                        //Write page attribute id
                        textwriter.WriteStartAttribute("id");
                        textwriter.WriteString(id);
                        textwriter.WriteEndAttribute();

                        //Write page attribute width
                        //textwriter.WriteStartAttribute("width");
                        //textwriter.WriteValue((int)(page.PageSheet.get_Cells("PageWidth").ResultIU * 96));
                        //textwriter.WriteEndAttribute();

                        //Write page attribute height
                        //textwriter.WriteStartAttribute("height");
                        //textwriter.WriteValue((int)(page.PageSheet.get_Cells("PageHeight").ResultIU * 96));
                        //textwriter.WriteEndAttribute();

                        //Write page attribute x
                        //textwriter.WriteStartAttribute("x");
                        //textwriter.WriteValue((int)(page.PageSheet.get_Cells("PinX").ResultIU * 96));
                        //textwriter.WriteEndAttribute();

                        //Write page attribute y
                        //textwriter.WriteStartAttribute("y");
                        //textwriter.WriteValue((int)(page.PageSheet.get_Cells("PinY").ResultIU * 96));
                        //textwriter.WriteEndAttribute();

                        //end shape
                        textwriter.WriteEndElement();

                        //在输出图片之前首先将覆盖在图片上方的图片名清除
                        shape.Text = "";
                        shape.Export(outputDir + "\\" + id + ".gif");
                    }
                }

                //Ends element page
                textwriter.WriteEndElement();

                //Ends the document
                textwriter.WriteEndDocument();
                textwriter.Flush();
                textwriter.Close();

                //程序退出时不保存对Visio的修改
                doc.Saved = true;

            }

            finally
            {
                doc.Close(); //关闭打开的文件
                app.Quit();  //退出Visio应用程序
            }
        }
    }
}
