using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace TestProject_framework.Reader
{
    public class xmlreaderclass
    {
       
        String xmlfilepath = @"C:\Users\al4104\source\repos\TestProject_framework\TestProject_framework\Reader\DataReader.xml";
        public void reader()
        {
            XmlDocument xmldoc = new XmlDocument();
            xmldoc.Load(xmlfilepath);
            xmldoc.Save(Console.Out);
        }

        public String DataReader(String Name)
        {
            XmlTextReader rd = new(xmlfilepath);
            String result = "";
            while (rd.Read())
            {
                // foreach (XmlNode node in runmode.ChildNodes)
                {
                    if (rd.NodeType == XmlNodeType.Element && rd.Name == Name)//Nodetype->gets current node type 
                    {
                        String Value = rd.ReadString();
                        Console.WriteLine(Name + ":" + Value);
                        result = Value;
                    }
                }
            }
            return result;
        }

    }
}
