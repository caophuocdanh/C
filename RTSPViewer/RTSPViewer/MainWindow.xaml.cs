using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using LibVLCSharp.Shared;

namespace RTSPViewer
{
    public partial class MainWindow : Window
    {
        private LibVLC _libVLC;
        private MediaPlayer _mediaPlayer;
        private Dictionary<string, List<string>> cameraLinks = new Dictionary<string, List<string>>();

        public MainWindow()
        {
            InitializeComponent();
            InitializeCameraLinks();
            cbBrand.SelectedIndex = 0; // Set default selection
            Core.Initialize();
            // Set the path to the native libraries directory if necessary
            string libVLCPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "libvlc"); // Adjust the folder name as necessary

            // Initialize LibVLC instance
            _libVLC = new LibVLC(libVLCPath);

            // Create MediaPlayer instance
            _mediaPlayer = new MediaPlayer(_libVLC);

            // Set the MediaPlayer to VideoView
            videoView.MediaPlayer = _mediaPlayer;

        }

        private void InitializeCameraLinks()
        {
            cameraLinks["Dahua"] = new List<string>
            {
                "rtsp://username:password@ipadd:port/cam/realmonitor?channel=1&subtype=0&unicast=true&proto=Onvif",
                "rtsp://username:password@ipadd:port/cam/realmonitor?channel=1&subtype=1"
            };

                    cameraLinks["HikVision"] = new List<string>
            {
                "rtsp://username:password@ipadd:port/video.h264",
                "rtsp://username:password@ipadd:port/0",
                "rtsp://username:password@ipadd:port/h264_stream",
                "rtsp://username:password@ipadd:port/live.dsp",
                "rtsp://username:password@ipadd:port/videomain",
                "rtsp://username:password@ipadd/Streaming/Channels/1"
            };

                    cameraLinks["Vivotek"] = new List<string>
            {
                "rtsp://username:password@ipadd/live.sdp"
            };

                    cameraLinks["Vantech"] = new List<string>
            {
                "rtsp://username:vantech123@ipadd:port/ch01/0"
            };

                    cameraLinks["KB Vision"] = new List<string>
            {
                "rtsp://username:password@ipadd:port/cam/realmonitor?channel=1&subtype=0&unicast=true&proto=Onvif"
            };
        }


        private async void btnViewCamera_Click(object sender, RoutedEventArgs e)
        {
            string brand = ((ComboBoxItem)cbBrand.SelectedItem).Content.ToString();
            string ipAddress = txtIpAddress.Text;
            if (!int.TryParse(txtPort.Text, out int port))
            {
                MessageBox.Show("Invalid port number.");
                return;
            }
            string username = txtUsername.Text;
            string password = txtPassword.Password;

            if (string.IsNullOrWhiteSpace(ipAddress) || string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
            {
                MessageBox.Show("Please fill in all fields.");
                return;
            }

            bool isCameraActive = false; // Variable to track if any RTSP link worked
            int i = 0;
            foreach (var link in cameraLinks[brand])
            {
                string rtspLink = link
                    .Replace("ipadd", ipAddress)
                    .Replace("port", port.ToString())
                    .Replace("username", username)
                    .Replace("password", password);

                if (CheckRTSPLink(rtspLink))
                {
                    try
                    {
                        _mediaPlayer.Play(new Media(_libVLC, new Uri(rtspLink)));
                        isCameraActive = true;
                        btnViewCamera.Content = "Loading....";
                        btnViewCamera.IsEnabled = false;
                        this.WindowState = WindowState.Maximized;
                        this.Topmost = true; // Optional: Keeps the window above others
                        break; // Exit loop once a working RTSP link is found
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error playing video: {ex.Message}");
                    }
                }
            }

            if (!isCameraActive)
            {
                MessageBox.Show("Camera is not active or RTSP is not enabled.");
            }
        }


        public static bool CheckRTSPLink(string rtspLink)
        {
            try
            {
                Uri uri = new Uri(rtspLink);

                // Create a TCP connection to the RTSP server
                using (TcpClient tcpClient = new TcpClient())
                {
                    tcpClient.Connect(uri.Host, uri.Port);

                    // Create a socket stream for the connection
                    using (NetworkStream stream = tcpClient.GetStream())
                    {
                        // Send an RTSP OPTIONS request to the server
                        string request = $"OPTIONS {rtspLink} RTSP/1.0\r\nCSeq: 1\r\n\r\n";
                        byte[] requestBytes = Encoding.ASCII.GetBytes(request);
                        stream.Write(requestBytes, 0, requestBytes.Length);

                        // Read the response from the server
                        byte[] responseBuffer = new byte[4096];
                        int bytesRead = stream.Read(responseBuffer, 0, responseBuffer.Length);

                        // Convert the response bytes to a string
                        string response = Encoding.ASCII.GetString(responseBuffer, 0, bytesRead);

                        // Check if the response contains "RTSP/1.0 200 OK"
                        if (response.Contains("RTSP/1.0 200 OK"))
                        {
                            // The RTSP link is valid and the server is responding
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking RTSP link: {ex.Message}");
            }

            // The RTSP link is not valid or the server is not responding
            return false;
        }


        protected override void OnClosed(EventArgs e)
        {
            // Clean up libVLC resources
            _mediaPlayer.Dispose();
            _libVLC.Dispose();
            base.OnClosed(e);
        }
    }
}
