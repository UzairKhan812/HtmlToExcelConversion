 public static class Extension
 {
     public static Task<FileStreamResult> TableToExcel<T>(List<T> list, string[] removeColumns, string name = "")
     {
         FileStreamResult fileStreamResult;
         DataTable table = list.ToDataTable();
         table.TableName = string.IsNullOrWhiteSpace(name) ? table.TableName : name;
         foreach (string ColName in removeColumns)
         {
             if (table.Columns.Contains(ColName))
                 table.Columns.Remove(ColName);
         }
         using (XLWorkbook wb = new XLWorkbook())
         {
             wb.Worksheets.Add(table);
             wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
             wb.Style.Font.Bold = true;
             MemoryStream stream = new MemoryStream();
             wb.SaveAs(stream);
             stream.Position = 0;
             fileStreamResult = new FileStreamResult(stream, "application/excel");
         }
         return Task.FromResult(fileStreamResult);
     }
     private static DataTable ToDataTable<T>(this List<T> items)
     {
         DataTable dataTable = new DataTable(typeof(T).Name);
         PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
         foreach (PropertyInfo prop in Props)
         {
             var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
             dataTable.Columns.Add(prop.Name, type);
         }
         foreach (T item in items)
         {
             var values = new object[Props.Length];
             for (int i = 0; i < Props.Length; i++)
             {
                 values[i] = Props[i].GetValue(item, null);
             }
             dataTable.Rows.Add(values);
         }
         return dataTable;
     }
 }
