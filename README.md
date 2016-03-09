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

Note: Export Process during; Excel title match class displayName data transfered is done
 
### Version

1.0.0

### Development

C#
  
