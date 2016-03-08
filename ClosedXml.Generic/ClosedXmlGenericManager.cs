using ClosedXML.Excel;
using JohnsonNet;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClosedXml.Generic
{
    public static class ClosedXmlGenericManager
    {
        /// <summary>
        /// A Sheet Export Class to Excel
        /// NOTE:Class DisplayName and Order Attiribute Required
        /// </summary> 
        /// <typeparam name="T"></typeparam>
        /// <param name="sheetName"></param>
        /// <param name="data"></param>
        /// <param name="fileName"></param>
        public static void Export<T>(string sheetName, IList<T> data, string filePath)
        {
            var workbook = new XLWorkbook();
            workbook.ToWorkSheet(data, sheetName);

            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// Multiple Sheet Export Class to Excel
        /// NOTE:Class DisplayName and Order Attiribute Required
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        public static void Export<T>(Dictionary<string, IList<T>> data, string filePath)
        {
            var workbook = new XLWorkbook();
            foreach (var item in data)
            {
                workbook.ToWorkSheet(item.Value, item.Key);
            }
            workbook.SaveAs(filePath);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static Dictionary<string, IList<T>> Import<T>(string filePath) where T : new()
        {
            Dictionary<string, IList<T>> sheetList = new Dictionary<string, IList<T>>();
            var workbook = new XLWorkbook(filePath);
            foreach (var item in workbook.Worksheets)
            {
                sheetList.Add(item.Name, ToEntity<T>(item));
            }
            return sheetList;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="workBook"></param>
        /// <param name="data"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        private static IXLWorksheet ToWorkSheet<T>(this XLWorkbook workBook, IList<T> data, string sheetName = "Sheet1")
        {
            var genericType = typeof(T);
            var workSheet = workBook.Worksheets.Add(sheetName);
            //ColumnProperty Info
            var columnList = GetColomnList(genericType);

            for (int i = 0; i < columnList.Count; i++)
            {
                var column = columnList[i];
                var cell = workSheet.Cell(1, i + 1);

                cell.Value = column.Attribute.Name;
                cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                cell.Style.Font.Bold = true;
            }

            // Create Rows 
            for (int rowIndex = 0; rowIndex < data.Count; rowIndex++)
            {
                var row = data[rowIndex];

                for (int columnIndex = 0; columnIndex < columnList.Count; columnIndex++)
                {
                    var column = columnList[columnIndex];
                    var cell = workSheet.Cell(rowIndex + 2, columnIndex + 1);
                    cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                    cell.Value = JohnsonManager.Convert.To(column.PropertyType, JohnsonManager.Reflection.GetPropertyValue(row, column.PropertyName));
                }
            }
            workSheet.Columns().AdjustToContents();
            return workSheet;
        }

        private static List<T> ToEntity<T>(this IXLWorksheet workSheet) where T : new()
        {
            var genericType = typeof(T);
            var columnList = GetColomnList(genericType);

            List<T> data = new List<T>();
            var rowList = workSheet.Rows().ToList();

            for (int i = 1; i < rowList.Count; i++)
            {
                var row = rowList[i];
                var instance = new T();

                for (int x = 0; x < columnList.Count; x++)
                {
                    var propertyName = columnList[x].PropertyName;
                    var property = genericType.GetProperty(propertyName);
                    var cell = row.Cell(x + 1);

                    property.SetValue(instance, JohnsonManager.Convert.To(property.PropertyType, cell.Value), null);
                }
                data.Add(instance);
            }
            return data;
        }

        private static List<InternalExcelHelperAttribute> GetColomnList(Type genericType)
        {
            var columnList = JohnsonManager.Reflection.GetPropertiesWithoutHidings(genericType)
                .Where(property => property.GetAttribute<DisplayAttribute>() != null)
                .Select(property =>
                {
                    return new InternalExcelHelperAttribute
                    {
                        Attribute = property.GetAttribute<DisplayAttribute>(),
                        PropertyName = property.Name,
                        PropertyType = property.PropertyType
                    };
                })
                .OrderBy(p => p.Attribute.GetOrder())
                .ToList();
            return columnList;
        }
    }
}
