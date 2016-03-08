# ClosedXml Generic Helper
 
Uzun zamandır ihtiyaç duyduğumuz generic excel import ve export işlemleri artık bu kütüphane ile çok kolay.

  - Hazır classlarımız üzerinden çekilen dataları excel'e import etme
  - Excel üzerindeki dataları classlarımız üzerinden export etme
  - 1 den fazla sheet oluşturarak içine farklı datalar aktarabilme
  
Bütün bunları bu library üzerinden çok kolay yapabilirsiniz.
 
### Installation

Nuget Package Manager Console

```sh
Install-Package ClosedXml.Generic
```
### Use of
Note: Excel import sırasında classlar üzerinde DisplayName boş bırakılırsa bu field altındaki datalar excel içerisine yüklenmez.

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

Note: Tam tersi olan export işleminde ise Excel içerisinde bulunan başlıklarla classlar üzerinde bulunan DisplayName alanları eşleştirilerek data aktarımı sağlanmaktadır.

 ### Version
1.0.0

### Development

C#
   
