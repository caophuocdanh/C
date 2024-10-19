using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using OfficeOpenXml;

namespace WarrantyChecker
{
    public partial class MainWindow : Window
    {
        private List<WarrantyInfo> warrantyInfos = new List<WarrantyInfo>();

        public MainWindow()
        {
            InitializeComponent();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Hoặc LicenseContext.Commercial nếu bạn có giấy phép thương mại
    }

    private async void BtnCheck_Click(object sender, RoutedEventArgs e)
        {
            var results = new List<WarrantyInfo>();

            foreach (var info in warrantyInfos)
            {
                var result = await CheckWarranty(info.SoSeria);
                if (result != null)
                {
                    result.STT = results.Count + 1; // Gán số thứ tự

                    // Đặt giá trị mặc định là mm/dd/yyyy
                    result.NgayXuat = DateTime.TryParse(result.NgayXuat, out var dateValue)
                        ? dateValue.ToString("MM/dd/yyyy")
                        : "N/A";

                    results.Add(result);
                }
            }
            dataGrid.ItemsSource = results;
        }
        private void CmbDateFormat_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Kiểm tra xem ItemsSource có hợp lệ không
            if (dataGrid.ItemsSource is List<WarrantyInfo> results && results.Count > 0)
            {
                var dateFormat = ((ComboBoxItem)cmbDateFormat.SelectedItem)?.Content.ToString();
                foreach (var result in results)
                {
                    if (DateTime.TryParse(result.NgayXuat, out var dateValue))
                    {
                        result.NgayXuat = dateFormat == "dd/MM/yyyy"
                            ? dateValue.ToString("dd/MM/yyyy")
                            : dateValue.ToString("MM/dd/yyyy");
                    }
                }
                dataGrid.Items.Refresh(); // Cập nhật DataGrid
            }
        }



        private async Task<WarrantyInfo> CheckWarranty(string seria)
        {
            using (HttpClient client = new HttpClient())
            {
                var response = await client.GetStringAsync($"https://app.dahua.vn:7778/Api.svc/Web/TraCuuBaoHanhTheoSeria?seria={seria}");
                var data = JsonSerializer.Deserialize<ApiResponse>(response);

                if (data != null && data.d != null)
                {
                    var results = JsonSerializer.Deserialize<List<WarrantyInfo>>(data.d);
                    if (results != null && results.Count > 0)
                    {
                        return results[0];
                    }
                    else
                    {
                        // Nếu danh sách trống, trả về thông tin với thông báo lỗi
                        return new WarrantyInfo
                        {
                            SoSeria = seria,
                            MaHangHoa = "Serial không đúng hoặc chưa đăng ký",
                            SoThangBaoHanh = 0,
                            NgayXuat = "",
                            SoNgayBaoHanhConLai = 0,
                            SoLanBaoHanh = 0
                        };
                    }
                }
            }
            return null;
        }


        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                Title = "Select a Text File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                var serials = File.ReadAllLines(openFileDialog.FileName);
                warrantyInfos.Clear(); // Xóa danh sách trước khi tải mới
                for (int i = 0; i < serials.Length; i++)
                {
                    warrantyInfos.Add(new WarrantyInfo { SoSeria = serials[i], STT = i + 1 }); // Gán số thứ tự
                }
                MessageBox.Show("Loaded serials!");
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveToExcel(warrantyInfos);
            MessageBox.Show("Data saved to Excel!");
        }

        private void SaveToExcel(List<WarrantyInfo> data)
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Warranty Data");
                worksheet.Cells[1, 1].Value = "STT";
                worksheet.Cells[1, 2].Value = "So Seria";
                worksheet.Cells[1, 3].Value = "Ma Hang Hoa";
                worksheet.Cells[1, 4].Value = "So Thang Bao Hanh";
                worksheet.Cells[1, 5].Value = "Ngay Xuat";
                worksheet.Cells[1, 6].Value = "So Ngay Bao Hanh Con Lai";
                worksheet.Cells[1, 7].Value = "So Lan Bao Hanh";

                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = data[i].STT; // Số thứ tự
                    worksheet.Cells[i + 2, 2].Value = data[i].SoSeria;
                    worksheet.Cells[i + 2, 3].Value = data[i].MaHangHoa;
                    worksheet.Cells[i + 2, 4].Value = data[i].SoThangBaoHanh;
                    worksheet.Cells[i + 2, 5].Value = data[i].NgayXuat;
                    worksheet.Cells[i + 2, 6].Value = data[i].SoNgayBaoHanhConLai;
                    worksheet.Cells[i + 2, 7].Value = data[i].SoLanBaoHanh;
                }

                var filePath = "WarrantyData.xlsx";
                File.WriteAllBytes(filePath, package.GetAsByteArray());
            }
        }
    }

    public class ApiResponse
    {
        public string d { get; set; }
    }

    public class WarrantyInfo
    {
        public int STT { get; set; } // Thêm thuộc tính STT
        public string SoSeria { get; set; }
        public string MaHangHoa { get; set; }
        public int SoThangBaoHanh { get; set; }
        public string NgayXuat { get; set; }
        public int SoNgayBaoHanhConLai { get; set; }
        public int SoLanBaoHanh { get; set; }
    }
}
