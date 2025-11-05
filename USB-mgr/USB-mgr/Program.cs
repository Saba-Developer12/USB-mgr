using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32.SafeHandles;

namespace RufusAnalog
{
    public partial class MainForm : Form
    {
        private ComboBox? usbComboBox;
        private TextBox? isoTextBox;
        private Button? browseButton;
        private Button? exitButton;
        private Button? startButton;
        private TextBox? logTextBox;
        private ProgressBar? progressBar;
        private bool isWorking = false;

        [StructLayout(LayoutKind.Sequential)]
        private struct STORAGE_DEVICE_NUMBER
        {
            public int DeviceType;
            public int DeviceNumber;
            public int PartitionNumber;
        }

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern SafeFileHandle CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr SecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool DeviceIoControl(
            SafeFileHandle hDevice,
            uint dwIoControlCode,
            IntPtr lpInBuffer,
            uint nInBufferSize,
            out STORAGE_DEVICE_NUMBER lpOutBuffer,
            int nOutBufferSize,
            out uint lpBytesReturned,
            IntPtr lpOverlapped
        );

        private const uint IOCTL_STORAGE_GET_DEVICE_NUMBER = 0x2D1080;
        private const uint GENERIC_READ = 0x80000000;
        private const uint FILE_SHARE_READ = 0x00000001;
        private const uint FILE_SHARE_WRITE = 0x00000002;
        private const uint OPEN_EXISTING = 3;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.usbComboBox = new ComboBox();
            this.isoTextBox = new TextBox();
            this.browseButton = new Button();
            this.exitButton = new Button();
            this.startButton = new Button();
            this.logTextBox = new TextBox();
            this.progressBar = new ProgressBar();

            // Form settings
            this.Text = "USB-mgr - Bootable USB Creator";
            this.Size = new Size(400, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // USB Devices Label
            Label usbLabel = new Label();
            usbLabel.Text = "USB Devices";
            usbLabel.Location = new Point(10, 10);
            this.Controls.Add(usbLabel);

            // USB ComboBox
            usbComboBox.Location = new Point(10, 30);
            usbComboBox.Size = new Size(360, 20);
            usbComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Controls.Add(usbComboBox);

            // ISO File Label
            Label isoLabel = new Label();
            isoLabel.Text = "ISO File";
            isoLabel.Location = new Point(10, 60);
            this.Controls.Add(isoLabel);

            // ISO TextBox
            isoTextBox.Location = new Point(10, 80);
            isoTextBox.Size = new Size(300, 20);
            this.Controls.Add(isoTextBox);

            // Browse Button
            browseButton.Text = "Browse";
            browseButton.Location = new Point(320, 80);
            browseButton.Size = new Size(50, 20);
            browseButton.Click += browseButton_Click;
            this.Controls.Add(browseButton);

            // Log TextBox
            logTextBox.Location = new Point(10, 110);
            logTextBox.Size = new Size(360, 80);
            logTextBox.Multiline = true;
            logTextBox.ReadOnly = true;
            logTextBox.ScrollBars = ScrollBars.Vertical;
            this.Controls.Add(logTextBox);

            // Exit Button
            exitButton.Text = "Exit";
            exitButton.Location = new Point(10, 200);
            exitButton.Click += exitButton_Click;
            this.Controls.Add(exitButton);

            // Start Button
            startButton.Text = "Start";
            startButton.Location = new Point(300, 200);
            startButton.Click += startButton_Click;
            this.Controls.Add(startButton);

            // ProgressBar
            progressBar.Location = new Point(0, this.ClientSize.Height - 20);
            progressBar.Size = new Size(this.ClientSize.Width, 20);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.ForeColor = Color.Green;
            progressBar.BackColor = Color.LightGray;
            progressBar.Value = 0;
            progressBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            this.Controls.Add(progressBar);

            this.Load += MainForm_Load;
            this.Resize += (s, e) => { progressBar!.Width = this.ClientSize.Width; };
        }

        private void MainForm_Load(object? sender, EventArgs e)
        {
            RefreshUsbList();
        }

        private void RefreshUsbList()
        {
            usbComboBox!.Items.Clear();
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType == DriveType.Removable && drive.IsReady)
                {
                    string item = $"{drive.Name} - {drive.VolumeLabel} ({Math.Round(drive.TotalSize / (1024.0 * 1024 * 1024), 1)} GB)";
                    usbComboBox.Items.Add(item);
                }
            }
        }

        private void browseButton_Click(object? sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "ISO files|*.iso";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                isoTextBox!.Text = ofd.FileName;
            }
        }

        private void exitButton_Click(object? sender, EventArgs e)
        {
            Application.Exit();
        }

        private async void startButton_Click(object? sender, EventArgs e)
        {
            if (isWorking) return;
            if (usbComboBox!.SelectedIndex == -1 || string.IsNullOrEmpty(isoTextBox!.Text))
            {
                MessageBox.Show("Please select a USB device and an ISO file.");
                return;
            }

            string selectedItem = usbComboBox.SelectedItem!.ToString()!;
            string driveLetter = selectedItem.Substring(0, 2); // e.g., "D:"
            string isoPath = isoTextBox.Text;

            if (MessageBox.Show($"This will erase all data on {driveLetter}. Continue?", "Warning", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }

            string? physicalDrive = GetPhysicalDriveFromLetter(driveLetter);
            if (physicalDrive == null)
            {
                MessageBox.Show("Cannot find physical drive.");
                return;
            }

            isWorking = true;
            startButton!.Enabled = false;
            browseButton!.Enabled = false;
            usbComboBox.Enabled = false;
            logTextBox!.Clear();
            progressBar!.Value = 0;

            await Task.Run(async () => await CreateBootableUsb(physicalDrive, isoPath));

            isWorking = false;
            startButton.Enabled = true;
            browseButton.Enabled = true;
            usbComboBox.Enabled = true;
            RefreshUsbList();
        }

        private async Task CreateBootableUsb(string physicalDrive, string isoPath)
        {
            try
            {
                AppendLog("Starting process...");

                // Step 1: Format the drive using diskpart
                AppendLog("Formatting USB drive...");
                UpdateProgress(10);
                string diskNumber = physicalDrive.Replace(@"\\.\PhysicalDrive", "");
                string tempScript = Path.GetTempFileName();
                using (StreamWriter sw = new StreamWriter(tempScript))
                {
                    sw.WriteLine($"select disk {diskNumber}");
                    sw.WriteLine("clean");
                    sw.WriteLine("create partition primary");
                    sw.WriteLine("select partition 1");
                    sw.WriteLine("format fs=fat32 quick");
                    sw.WriteLine("active");
                    sw.WriteLine("assign");
                    sw.WriteLine("exit");
                }

                string diskpartOutput = await RunProcessAsync("diskpart.exe", "/s " + tempScript, true);
                AppendLog(diskpartOutput);
                File.Delete(tempScript);
                UpdateProgress(30);

                // Wait for USB drive to be ready
                AppendLog("Waiting for USB drive to be detected...");
                string? usbLetter = null;
                int attempts = 0;
                while (string.IsNullOrEmpty(usbLetter) && attempts < 10)
                {
                    await Task.Delay(1000);
                    usbLetter = GetDriveLetterFromPhysical(physicalDrive);
                    attempts++;
                }
                if (string.IsNullOrEmpty(usbLetter))
                {
                    throw new Exception("Cannot find USB drive letter after formatting.");
                }
                AppendLog($"USB drive detected at {usbLetter}");

                // Step 2: Mount ISO
                AppendLog("Mounting ISO...");
                string mountCommand = $"-Command \"$di = Mount-DiskImage -ImagePath '{isoPath.Replace("\\", "\\\\")}' -PassThru; (Get-Volume -DiskImage $di).DriveLetter\"";
                string mountOutput = await RunProcessAsync("powershell.exe", mountCommand, false);
                AppendLog("Mount output: " + mountOutput);
                string mountedLetterStr = mountOutput.Trim();
                if (string.IsNullOrEmpty(mountedLetterStr) || mountedLetterStr.Contains("Error"))
                {
                    throw new Exception("Cannot detect mounted ISO drive letter: " + mountOutput);
                }
                string mountedLetter = mountedLetterStr + ":\\";
                AppendLog($"ISO mounted at {mountedLetter}");
                UpdateProgress(40);

                // Step 3: Copy files
                AppendLog("Copying files to USB...");
                CopyDirectoryWithProgress(mountedLetter, usbLetter + "\\");
                UpdateProgress(80);

                // Step 4: Dismount ISO
                AppendLog("Dismounting ISO...");
                string dismountOutput = await RunProcessAsync("powershell.exe", $"-Command Dismount-DiskImage -ImagePath '{isoPath.Replace("\\", "\\\\")}'", false);
                AppendLog(dismountOutput);
                UpdateProgress(100);

                AppendLog("Bootable USB created successfully!");
                InvokeMessageBox("Success", "Bootable USB created successfully!");
            }
            catch (Exception ex)
            {
                AppendLog("Error: " + ex.Message);
                UpdateProgress(0);
                InvokeMessageBox("Error", "Failed to create bootable USB: " + ex.Message);
            }
        }

        private void CopyDirectoryWithProgress(string sourceDir, string targetDir)
        {
            var files = Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories);
            int totalFiles = files.Length;
            int copied = 0;

            foreach (string file in files)
            {
                string relativePath = file.Substring(sourceDir.Length).TrimStart('\\');
                string destFile = Path.Combine(targetDir, relativePath);
                string? destDir = Path.GetDirectoryName(destFile);
                if (!string.IsNullOrEmpty(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }
                File.Copy(file, destFile, true);
                copied++;
                int progress = 40 + (int)(40 * (copied / (double)totalFiles)); // From 40% to 80%
                UpdateProgress(progress);
            }
        }

        private List<string> GetCdDrives()
        {
            return DriveInfo.GetDrives()
                .Where(d => d.DriveType == DriveType.CDRom && d.IsReady)
                .Select(d => d.Name)
                .ToList();
        }

        private async Task<string> RunProcessAsync(string fileName, string arguments, bool runAsAdmin)
        {
            Process process = new Process();
            process.StartInfo.FileName = fileName;
            process.StartInfo.Arguments = arguments;
            process.StartInfo.UseShellExecute = runAsAdmin;
            process.StartInfo.RedirectStandardOutput = !runAsAdmin;
            process.StartInfo.RedirectStandardError = !runAsAdmin;
            process.StartInfo.CreateNoWindow = true;
            if (runAsAdmin) process.StartInfo.Verb = "runas";

            process.Start();
            string output = "";
            if (!runAsAdmin)
            {
                string stdOut = await process.StandardOutput.ReadToEndAsync();
                string stdErr = await process.StandardError.ReadToEndAsync();
                output = stdOut + (string.IsNullOrEmpty(stdErr) ? "" : "\nError: " + stdErr);
            }
            await Task.Run(() => process.WaitForExit());
            return output;
        }

        private void AppendLog(string message)
        {
            this.Invoke(new Action(() => 
            {
                logTextBox!.AppendText(message + "\r\n");
                logTextBox.ScrollToCaret();
            }));
        }

        private void UpdateProgress(int value)
        {
            this.Invoke(new Action(() => progressBar!.Value = Math.Min(value, 100)));
        }

        private void InvokeMessageBox(string title, string message)
        {
            this.Invoke(new Action(() => MessageBox.Show(message, title)));
        }

        private int GetDiskNumberFromLetter(string letter)
        {
            string volPath = @"\\.\" + letter.TrimEnd(':') + ":";
            using (SafeFileHandle handle = CreateFile(volPath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero))
            {
                if (handle.IsInvalid)
                {
                    AppendLog("CreateFile failed: " + Marshal.GetLastWin32Error());
                    return -1;
                }

                STORAGE_DEVICE_NUMBER sdn;
                uint bytesReturned;
                if (DeviceIoControl(handle, IOCTL_STORAGE_GET_DEVICE_NUMBER, IntPtr.Zero, 0, out sdn, Marshal.SizeOf<STORAGE_DEVICE_NUMBER>(), out bytesReturned, IntPtr.Zero))
                {
                    return sdn.DeviceNumber;
                }
                else
                {
                    AppendLog("DeviceIoControl failed: " + Marshal.GetLastWin32Error());
                    return -1;
                }
            }
        }

        private string? GetPhysicalDriveFromLetter(string letter)
        {
            int diskNum = GetDiskNumberFromLetter(letter);
            if (diskNum == -1)
            {
                return null;
            }
            return @"\\.\PhysicalDrive" + diskNum;
        }

        private string? GetDriveLetterFromPhysical(string physical)
        {
            string diskNumStr = physical.Replace(@"\\.\PhysicalDrive", "");
            if (!int.TryParse(diskNumStr, out int targetDiskNum))
            {
                return null;
            }

            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType == DriveType.Removable && drive.IsReady)
                {
                    string dl = drive.Name.Substring(0, 2);
                    int dn = GetDiskNumberFromLetter(dl);
                    if (dn == targetDiskNum)
                    {
                        return dl;
                    }
                }
            }
            return null;
        }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}