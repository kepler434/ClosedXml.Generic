# ClosedXml Generic Helper
 
Generic Excel import and export processes are too easy by using this library. 

- Importing the data from existing classes
- Exporting the data from existing classes
- More then one excel sheet create 

You can accomplish all these tasks easily by using ClosedXml.Generic
 
### Installation

Nuget Package Manager Console

```sh
Install-Package ClosedXml.Generic
```
### Use of
Note: If DisplayName field is empty in anyone of the classes, data related to this field can not be loaded to excel file.

##Import 
```sh
....
using System.ComponentModel.DataAnnotations;
..... 
        static void Main(string[] args)
        {
            ClosedXmlGenericManager.Import<MyClassName>("C:\MyExcel");
        } 

        public class MyClassName
        {
            [Display(Name = "Baslik", Order = 1)]
            public string Name { get; set; }

            [Display(Name = "Kayıt Numarası", Order = 2)]
            public int ID { get; set; }
 
            public string LastName { get; set; }
        }
```
##Export 
```sh
....
using System.ComponentModel.DataAnnotations;
..... 
        static void Main(string[] args)
        {
            Dictionary<string, IList<MyClassName>> data = new Dictionary<string, IList<MyClassName>>();
            
            data.Add("Sheet1", new List<MyClassNameMyClassName>()
            {
              new MyClassName() {ID = 1, Name ="Soner", LastName = "Kavlak"},
              new MyClassName() {ID = 2, Name ="John", LastName = "Durk"},
            });
            
            ClosedXmlGenericManager.Import<MyClassName>(data,"C:\MyExcel");
        }  
```

Note: Data transfer is completed by matching excel titles and class DisplayName's during export process



##New Version Dictionary List Object

```sh
....
using System.ComponentModel.DataAnnotations;
..... 
        static void Main(string[] args)
        {
            List<Dictionary<string, object>> dictionary = new List<Dictionary<string, object>>();
            for (int i = 0; i < 9; i++)
            {
                Dictionary<string, object> dict = new Dictionary<string, object>();
                dict.Add("cat", i + 5);
                dict.Add("dog", "deneme"+i);
                dict.Add("llama", i + 2.5);
                dict.Add("iguana", i + 8);
                dictionary.Add(dict);
           }
           ClosedXml.Generic.ClosedXmlGenericManager.Export("Sheet1", dictionary, "C:\MyExcel"); 
        }  
```
 
 
### Version

1.0.4

### Development

C#
  
