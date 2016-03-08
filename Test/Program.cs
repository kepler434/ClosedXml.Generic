using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks; 
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Dapper;

namespace Test
{
    class Program
    { 
        static void Main(string[] args)
        {
            Export();
            //Import();
        }

        public static void PivotImport()
        { 
            
        }

        public static void Import()
        {
            //ExcelHelper.ExcelOperation.Import<ListValue>();
        }

        public static void Export()
        {
            Dictionary<string, IList<ListValue>> data = new Dictionary<string, IList<ListValue>>();

            data.Add("Sheet1", new List<ListValue>()
            {  
              new ListValue() {ID = 1, Name ="Soner", LastName = "Kavlak"},
              new ListValue() {ID = 2, Name ="Emrah", LastName = "Pınar"}, 
            });

            data.Add("Sheet2", new List<ListValue>()
            {  
              new ListValue() {ID = 4, Name ="Mert"},
              new ListValue() {ID = 5, Name ="Burak"},
            });

            ClosedXml.Generic.ClosedXmlGenericManager.Export<ListValue>(data, ""); 
        }

        public class ListValue
        {
            [Display(Name = "Baslik", Order = 1)]
            public string Name { get; set; }

            [Display(Name = "Kayıt Numarası")]
            public int ID { get; set; }


            public string LastName { get; set; }
        }
    }
}
