using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml.Generic
{
    public class InternalExcelHelperAttribute
    {
        public DisplayAttribute Attribute { get; set; }
        public string PropertyName { get; set; }
        public Type PropertyType { get; set; }
    }
}
