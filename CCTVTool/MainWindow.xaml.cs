using System;
using System.Collections.ObjectModel;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32; // Thêm thư viện để dùng OpenFileDialog
using Microsoft.Office.Interop.Excel;
using System.Text;
using System.Net;
using System.Collections.Generic;

namespace CCTVTool
{
    public partial class MainWindow : System.Windows.Window
    {
        public ObservableCollection<CameraInfo> Cameras { get; set; } = new ObservableCollection<CameraInfo>();

        public MainWindow()
        {
            InitializeComponent();
            cameraDataGrid.ItemsSource = Cameras; // Gán vào DataGrid
            this.Title = "CCTV Tool @Danh";
        }

        // Camera info model
        public class CameraInfo
        {
            public int STT { get; set; }
            public string IP { get; set; }
            public string User { get; set; }
            public string Password { get; set; }
            public string Model { get; set; }
            public string Status { get; set; }
        }

        // Hàm xử lý khi nhấn nút tạo file Excel mẫu
        private void CreateSampleExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Tạo ứng dụng Excel
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Add();
                Worksheet worksheet = workbook.Sheets[1];

                // Tạo tiêu đề cột
                worksheet.Cells[1, 1] = "IP";
                worksheet.Cells[1, 2] = "User";
                worksheet.Cells[1, 3] = "Password";
                worksheet.Cells[1, 4] = "Model";

                // Tạo nội dung mẫu
                worksheet.Cells[2, 1] = "192.168.1.10";
                worksheet.Cells[2, 2] = "admin";
                worksheet.Cells[2, 3] = "admin123";
                worksheet.Cells[2, 4] = "Dahua";

                worksheet.Cells[3, 1] = "192.168.1.11";
                worksheet.Cells[3, 2] = "admin";
                worksheet.Cells[3, 3] = "admin123";
                worksheet.Cells[3, 4] = "Hikvision";

                worksheet.Cells[4, 1] = "192.168.1.12";
                worksheet.Cells[4, 2] = "admin";
                worksheet.Cells[4, 3] = "admin123";
                worksheet.Cells[4, 4] = "ONVIF";

                // Hiển thị hộp thoại lưu file
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Files|*.xlsx",
                    Title = "Lưu file mẫu",
                    FileName = "CCTV_Sample.xlsx"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                    MessageBox.Show("Tạo file mẫu thành công!", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                }

                // Đóng và thoát ứng dụng Excel
                workbook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi tạo file Excel: {ex.Message}");
            }
        }

        // Load cameras from Excel file
        private void LoadCamerasFromExcel(string filePath)
        {
            try
            {
                // Tạo ứng dụng Excel
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Workbooks.Open(filePath);
                Worksheet worksheet = workbook.Sheets[1];
                Range range = worksheet.UsedRange;

                Cameras.Clear(); // Xóa dữ liệu cũ trước khi nạp mới

                for (int row = 2; row <= range.Rows.Count; row++) // Giả sử hàng đầu tiên là tiêu đề
                {
                    Cameras.Add(new CameraInfo
                    {
                        STT = row - 1,
                        IP = (range.Cells[row, 1] as Range).Value2.ToString(),
                        User = (range.Cells[row, 2] as Range).Value2.ToString(),
                        Password = (range.Cells[row, 3] as Range).Value2.ToString(),
                        Model = (range.Cells[row, 4] as Range).Value2.ToString(),
                        Status = "Unknown" // Lúc đầu chưa kiểm tra kết nối
                    });
                }

                workbook.Close(false);
                excelApp.Quit();

                // Sau khi load dữ liệu từ file Excel, kiểm tra trạng thái của từng camera
                CheckAllCamerasStatus(); // Kiểm tra tất cả camera và cập nhật DataGrid
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Lỗi khi đọc file Excel: {ex.Message}");
            }
        }



        // Xử lý sự kiện Browse button click
        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files|*.xlsx;*.xls",
                Title = "Chọn file Excel"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                LoadCamerasFromExcel(openFileDialog.FileName); // Gọi hàm để load dữ liệu và kiểm tra trạng thái
            }
        }


        private async void CheckAllCamerasStatus()
        {
            foreach (var camera in Cameras)
            {
                await CheckCameraStatus(camera); // Kiểm tra từng camera
            }
            cameraDataGrid.Items.Refresh(); // Cập nhật lại DataGrid
        }


        private async Task CheckCameraStatus(CameraInfo camera)
        {
            using (HttpClient client = new HttpClient())
            {
                try
                {
                    // Gửi yêu cầu GET tới địa chỉ IP của camera
                    HttpResponseMessage response = await client.GetAsync($"http://{camera.IP}");

                    // Nếu phản hồi thành công (status code 200), camera online
                    if (response.IsSuccessStatusCode)
                    {
                        camera.Status = "Online";
                    }
                    else
                    {
                        camera.Status = "Offline";
                    }
                }
                catch
                {
                    // Nếu có lỗi xảy ra (không kết nối được), camera offline
                    camera.Status = "Offline";
                }
            }
        }


        // Gửi yêu cầu reboot camera
        private async Task RebootCamera(CameraInfo camera)
        {
            try
            {
                // Thiết lập HttpClientHandler với Digest Authentication cho Hikvision
                var handler = new HttpClientHandler
                {
                    Credentials = new NetworkCredential(camera.User, camera.Password),
                    PreAuthenticate = true
                };

                using (HttpClient client = new HttpClient(handler))
                {
                    // Kiểm tra model của camera và chọn URL reboot phù hợp
                    string rebootUrl = "";
                    HttpMethod method = new HttpMethod("POST"); // Mặc định cho Dahua và ONVIF
                    if (camera.Model.Contains("Dahua"))
                    {
                        rebootUrl = $"http://{camera.IP}/cgi-bin/magicBox.cgi?action=reboot";
                    }
                    else if (camera.Model.Contains("Hikvision"))
                    {
                        rebootUrl = $"http://{camera.IP}/ISAPI/System/reboot";
                        method = new HttpMethod("PUT"); // Hikvision yêu cầu sử dụng PUT
                    }
                    else if (camera.Model.Contains("ONVIF"))
                    {
                        // ONVIF thường yêu cầu SOAP, nhưng tạm sử dụng HTTP cho reboot nếu hỗ trợ
                        rebootUrl = $"http://{camera.IP}/onvif/device_service";
                    }
                    else
                    {
                        camera.Status = "Model not recognized for reboot.";
                        return;
                    }

                    // Thực hiện yêu cầu reboot
                    HttpRequestMessage request = new HttpRequestMessage(method, rebootUrl);
                    HttpResponseMessage response = await client.SendAsync(request);

                    if (response.IsSuccessStatusCode)
                    {
                        camera.Status = "Rebooted successfully.";
                    }
                    else
                    {
                        camera.Status = $"Failed: {response.StatusCode}";
                    }
                }
            }
            catch (Exception ex)
            {
                camera.Status = $"Error: {ex.Message}";
            }
        }

        private async Task RebootAllCameras(List<CameraInfo> cameras)
        {
            foreach (var camera in cameras)
            {
                await RebootCamera(camera);
                // Cập nhật trạng thái vào DataGrid hoặc giao diện
            }
        }

        // Nút bấm reboot
        private async void RebootButton_Click(object sender, RoutedEventArgs e)
        {
            foreach (var camera in Cameras)
            {
                if (camera.Status == "Online")
                {
                    await RebootCamera(camera);
                }
            }
            cameraDataGrid.Items.Refresh(); // Cập nhật lại DataGrid
        }
    }
}
