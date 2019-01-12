using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompareExcelToXMLConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Program Start.");

            
            List<string> keyList = new List<string>();
            keyList.Add("JOBNAME");
            keyList.Add("MEMNAME");
            keyList.Add("DESCRIPTION");
            keyList.Add("RUN_AS");
            keyList.Add("TASKTYPE");
            keyList.Add("NODEID");
            keyList.Add("MEMLIB");
            keyList.Add("TIMEFROM");


            XMLHandling xmlFile = new XMLHandling();
            var xmlData = xmlFile.ReadXMLData(ConfigurationManager.AppSettings["XMLFilePath"], keyList);
            Console.ReadLine();
        }
    }
}
