using ExcelDataReader;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Studio.DreamRoom.OrderSheetConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string currentFilePath = "";

        private SheetData? sheetData = null;

        public MainWindow()
        {
            InitializeComponent();

            var version = Assembly.GetExecutingAssembly().GetName().Version;

            this.Title = $"订单表格转换器 ({version})";

            System.Windows.Forms.Application.EnableVisualStyles();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        private void DropTargetPanel_Drop(object sender, DragEventArgs e)
        {
            ExcelLogo.Opacity = 1;

            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    HandleFileOpen(files[0]);
                }
            }
        }

        private void HandleFileOpen(string file)
        {
            if (File.Exists(file) && Path.GetExtension(file).ToLower() == ".xls")
            {
                using (var stream = File.Open(file, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet();

                        var table = result.Tables.OfType<DataTable>().FirstOrDefault(x => x.TableName == "接龙列表(行排不合并)");
                        if (table != null)
                        {
                            FilePathText.Text = file;

                            currentFilePath = file;

                            var sheetData = SheetParser.Parse(table);
                            MainView.IsEnabled = true;
                            ShowSheetData(sheetData);
                            this.sheetData = sheetData;
                        }
                        else
                        {
                            Utils.ShowErrorMessage("无法识别的 Excel 文件。未在该文件中找到名为“接龙列表(行排不合并)”的表格。", "打开文件");
                        }
                    }
                }
            }
            else
            {
                Utils.ShowErrorMessage("文件不存在或者格式错误。请确认您选择的文件为有效的 Excel 97-2003 Workbook (*.xls)。", "打开文件");
            }
        }

        private void DropTargetPanel_DragEnter(object sender, DragEventArgs e)
        {
            ExcelLogo.Opacity = 0.85;
        }

        private void DropTargetPanel_DragLeave(object sender, DragEventArgs e)
        {
            ExcelLogo.Opacity = 1;
        }

        private void ShowSheetData(SheetData sheetData)
        {
            ProductsCountText.Text = sheetData.Products.Count.ToString();
            BuyersCountText.Text = sheetData.Buyers.Count.ToString();
            OriginalOrdersCountText.Text = sheetData.OriginalOrdersCount.ToString();
            OrdersCountText.Text = sheetData.OrdersCount.ToString();

            BuyersTree.Header = $"买家 ({sheetData.Buyers.Count})";
            ProductsTree.Header = $"产品 ({sheetData.Products.Count})";

            BuyersTree.ItemsSource = sheetData.Buyers;
            ProductsTree.ItemsSource = sheetData.Products;
        }

        private void ExcelLogo_Click(object sender, MouseButtonEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel 97-2003 Workbook|*.xls";
            dialog.DefaultExt = ".xls";

            var result = dialog.ShowDialog();
            if (result == true)
            {
                HandleFileOpen(dialog.FileName);
            }
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog();
            dialog.Filter = "Excel Workbook|*.xlsx";
            dialog.DefaultExt = ".xlsx";
            dialog.FileName = $"{Path.GetFileNameWithoutExtension(currentFilePath)}-converted";

            var result = dialog.ShowDialog();
            if (result == true)
            {
                using (var p = new ExcelPackage())
                {
                    if (sheetData != null) {
                        try
                        {
                            SheetGenerator.Export((SheetData)sheetData, dialog.FileName, currentFilePath);


                            if (File.Exists(dialog.FileName))
                            {
                                Utils.ShowSuccessMessage($"文件导出成功。", "导出文件");
                            }
                        }
                        catch (Exception ex)
                        {
                            Utils.ShowErrorMessage($"导出文件时发生错误，请确认同名文件未被 Excel 打开中。\n\nException: {ex.Message}", "导出文件");
                        }
                    }
                }
            }
        }

        private void AboutText_Click(object sender, MouseButtonEventArgs e)
        {
            Process.Start("explorer.exe", "https://dreamroom.studio");
        }

        private void FilePathText_Click(object sender, MouseButtonEventArgs e)
        {
            var path = ((TextBlock)sender).Text;
            if (path != null && File.Exists(path))
            {
                var args = $"/select, \"{path}\"";
                Process.Start("explorer.exe", args);
            }
        }
    }
}
