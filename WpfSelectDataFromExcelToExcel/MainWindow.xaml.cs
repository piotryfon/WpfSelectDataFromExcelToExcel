using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfSelectDataFromExcelToExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var products = ImportExcel<Product>(@"C:\excel\products.xlsx", "products");
            var selectedProd = products.Where(p => p.Name != "" && p.Price > 3);

            var wb = new XLWorkbook();

            var ws = wb.Worksheets.Add("selected-products");
            ws.Cell(1, 1).Value = "Product";
            ws.Cell(1, 2).Value = "Price";
            ws.Cell(1, 3).Value = "Units";
            ws.Cell(2, 1).InsertData(selectedProd);
            var col1 = ws.Column("A");
            col1.Width = 20;
            //col1.Style.Fill.BackgroundColor = XLColor.Orange;
            wb.SaveAs($"c:\\excel\\selected_prod_{DateTime.Now.ToString("MM-dd-yyyy")}.xlsx");
        }
        public List<T> ImportExcel<T>(string excelFilePath, string sheetName)
        {
            List<T> list = new List<T>();
            Type typeOfObject = typeof(T);
            using (IXLWorkbook workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheets.Where(w => w.Name == sheetName).First();
                var properties = typeOfObject.GetProperties();
                var columns = worksheet.FirstRow().Cells().Select((v, i) => new { Value = v.Value, Index = i + 1 });
                try
                {
                    foreach (IXLRow row in worksheet.RowsUsed().Skip(1))
                    {
                        T obj = (T)Activator.CreateInstance(typeOfObject);
                        foreach (var prop in properties)
                        {
                            int colIndex = columns.SingleOrDefault(c => c.Value.ToString() == prop.Name.ToString()).Index;
                            var val = row.Cell(colIndex).Value;
                            var type = prop.PropertyType;
                            prop.SetValue(obj, Convert.ChangeType(val, type));
                        }
                        list.Add(obj);
                    }
                    Console.WriteLine("Operacja zakończona powodzeniem.");
                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message);
                }
            }
            return list;
        }
    }
}
