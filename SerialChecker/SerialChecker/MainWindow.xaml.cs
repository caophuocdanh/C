using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;
using System.Collections.Generic;

namespace SerialChecker
{
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

    public class ApiResponse
    {
        public string D { get; set; }
    }

    public partial class MainWindow : Window
    {
        private ObservableCollection<WarrantyInfo> results = new ObservableCollection<WarrantyInfo>();
        private readonly HttpClient httpClient = new HttpClient();
        private List<string> warrantyInfos = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            listView.ItemsSource = results; // Gán nguồn dữ liệu
        }

        private async Task<WarrantyInfo> CheckWarranty(string serial)
        {
            string url = $"https://app.dahua.vn:7778/Api.svc/Web/TraCuuBaoHanhTheoSeria?seria={serial}";
            var response = await httpClient.GetStringAsync(url);
            var apiResponse = JsonConvert.DeserializeObject<ApiResponse>(response);

            if (apiResponse.D.Length == 0 || apiResponse.D == "[]")
            {
                return new WarrantyInfo
                {
                    SoSeria = serial,
                    MaHangHoa = "Serial không tồn tại hoặc chưa đăng ký"
                };
            }

            var warrantyData = JsonConvert.DeserializeObject<List<WarrantyInfo>>(apiResponse.D);
            return warrantyData[0]; // Lấy thông tin đầu tiên từ danh sách
        }

        private async void BtnCheck_Click(object sender, RoutedEventArgs e)
        {
            results.Clear(); // Xóa kết quả trước khi kiểm tra

            int index = 1; // Khởi tạo chỉ số STT
            foreach (var info in warrantyInfos)
            {
                var result = await CheckWarranty(info);
                if (result != null)
                {
                    result.STT = index++; // Gán số thứ tự
                    results.Add(result); // Thêm vào ObservableCollection
                }
            }
        }


        private void BtnLoad_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Text files (*.txt)|*.txt"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                var lines = File.ReadAllLines(openFileDialog.FileName);
                warrantyInfos = new List<string>(lines); // Tải số seri từ file
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Warranty Info");
                    worksheet.Cells[1, 1].Value = "STT"; // Thêm tiêu đề cho STT
                    worksheet.Cells[1, 2].Value = "So Seria";
                    worksheet.Cells[1, 3].Value = "Ma Hang Hoa";
                    worksheet.Cells[1, 4].Value = "So Thang Bao Hanh";
                    worksheet.Cells[1, 5].Value = "Ngay Xuat";
                    worksheet.Cells[1, 6].Value = "So Ngay Bao Hanh Con Lai";
                    worksheet.Cells[1, 7].Value = "So Lan Bao Hanh";

                    for (int i = 0; i < results.Count; i++)
                    {
                        worksheet.Cells[i + 2, 1].Value = results[i].STT; // Ghi STT vào file Excel
                        worksheet.Cells[i + 2, 2].Value = results[i].SoSeria;
                        worksheet.Cells[i + 2, 3].Value = results[i].MaHangHoa;
                        worksheet.Cells[i + 2, 4].Value = results[i].SoThangBaoHanh;
                        worksheet.Cells[i + 2, 5].Value = results[i].NgayXuat;
                        worksheet.Cells[i + 2, 6].Value = results[i].SoNgayBaoHanhConLai;
                        worksheet.Cells[i + 2, 7].Value = results[i].SoLanBaoHanh;
                    }

                    package.SaveAs(new FileInfo(saveFileDialog.FileName));
                }
            }
        }

    }
}
