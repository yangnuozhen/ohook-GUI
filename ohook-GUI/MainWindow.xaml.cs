using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
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
using System.Reflection;
using static System.Net.Mime.MediaTypeNames;

namespace ohook_GUI
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if(O365HomePrem.IsChecked == true)
            {
                Crack(0);
            }else if(O365ProPlus.IsChecked == true)
            {
                Crack(1);
            }
            
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="productType">值为 0 为 O365HomePrem ; 值为 1 则为 O365ProPlus</param>
        private async Task Crack(int productType)
        {
            if (productType != 0 && productType != 1)
            {
                throw new Exception("Unexpected productType input!");
            }
            await UpdateProgressBar(14.29);
            await UpdateStatusAsync("释放 sppc64.dll..");
            string tempFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "sppc64.dll");
            string destinationFilePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Office", "root", "vfs", "System", "sppc.dll");
            string[] productBlankKey = { "NBBBB-BBBBB-BBBBB-BBBBG-234RY", "NBBBB-BBBBB-BBBBB-BBBCF-PPK9C" };
            byte[] fileBytes = GetEmbeddedFileBytes("sppc64.dll");
            await UpdateProgressBar(28.58);
            await Task.Run(() => File.WriteAllBytes(tempFilePath, fileBytes));
            await UpdateProgressBar(42.87);
            await UpdateStatusAsync("正在创建符号链接...");
            await Task.Run(() => ExecuteCommand("mklink \"%programfiles%\\Microsoft Office\\root\\vfs\\System\\sppcs.dll\" \"%windir%\\System32\\sppc.dll\""));
            await UpdateProgressBar(57.16);
            await UpdateStatusAsync("将 sppc64.dll 复制到Office 中..");
            await Task.Run(() => ExecuteCommand($"copy /y \"{tempFilePath}\" \"{destinationFilePath}\""));
            await UpdateProgressBar(71.45);
            await UpdateStatusAsync("正在写入Host....");
            await Task.Run(() => ExecuteCommand("echo 0.0.0.0 ols.officeapps.live.com >> %windir%\\System32\\drivers\\etc\\hosts"));
            await UpdateProgressBar(85.74);
            await UpdateStatusAsync("注册产品...");
            await Task.Run(() => ExecuteCommand($"slmgr -ipk {productBlankKey[productType]}"));
            await UpdateProgressBar(99);
            await UpdateStatusAsync("成功!");
            await UpdateProgressBar(100);
        }

        private Task UpdateStatusAsync(string text)
        {
            return Task.Run(() =>
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    Status.Text = text;
                });
            });
        }
        private Task UpdateProgressBar(double Value)
        {
            return Task.Run(() =>
            {
                System.Windows.Application.Current.Dispatcher.Invoke(() =>
                {
                    progress.Value = Value;
                });
            });
        }
        static void ExecuteCommand(string command)
        {
            ProcessStartInfo processInfo = new ProcessStartInfo("cmd.exe", "/c " + command);
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;

            Process process = new Process();
            process.StartInfo = processInfo;
            process.Start();
            process.WaitForExit();
        }
        static byte[] GetEmbeddedFileBytes(string fileName)
        {
            using (var stream = Assembly.GetEntryAssembly().GetManifestResourceStream("ohook_GUI.res." + fileName))
            {
                if (stream == null)
                {
                    throw new Exception($"Embedded resource '{fileName}' not found.");
                }

                byte[] buffer = new byte[stream.Length];
                stream.Read(buffer, 0, buffer.Length);
                return buffer;
            }
        }


    }
}
