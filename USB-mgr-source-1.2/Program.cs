using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using Microsoft.Win32.SafeHandles;

namespace RufusAnalog
{
    public static class Localization
    {
        public static string FormTitle { get; private set; } = "USB-mgr";
        public static string UsbDevices { get; private set; } = "USB Devices";
        public static string IsoFile { get; private set; } = "ISO File";
        public static string Browse { get; private set; } = "Browse";
        public static string Exit { get; private set; } = "Exit";
        public static string Start { get; private set; } = "Start";
        public static string SelectUsbAndIso { get; private set; } = "Please select a USB device and an ISO file.";
        public static string EraseWarning { get; private set; } = "This will erase all data on {0}. Continue?";
        public static string Warning { get; private set; } = "Warning";
        public static string NoPhysicalDrive { get; private set; } = "Cannot find physical drive.";
        public static string StartingProcess { get; private set; } = "Starting process...";
        public static string FormattingUsb { get; private set; } = "Cleaning USB drive (deleting files only)...";
        public static string MountingIso { get; private set; } = "Mounting ISO...";
        public static string IsoMountedAt { get; private set; } = "ISO mounted at {0}";
        public static string CopyingFiles { get; private set; } = "Copying files to USB...";
        public static string DismountingIso { get; private set; } = "Dismounting ISO...";
        public static string BootableSuccess { get; private set; } = "Bootable USB created successfully!";
        public static string Success { get; private set; } = "Success";
        public static string ErrorFormat { get; private set; } = "Error: {0}";
        public static string Error { get; private set; } = "Error";
        public static string WaitingUsb { get; private set; } = "Waiting for USB drive to be detected...";
        public static string UsbDetectedAt { get; private set; } = "USB drive detected at {0}";
        public static string NoIsoLetter { get; private set; } = "Cannot detect mounted ISO drive letter: {0}";
        public static string NoUsbLetterAfterFormat { get; private set; } = "Cannot find USB drive letter.";
        public static string MountOutput { get; private set; } = "Mount output: {0}";
        public static string CreateFileFailed { get; private set; } = "CreateFile failed: {0}";
        public static string DeviceIoControlFailed { get; private set; } = "DeviceIoControl failed: {0}";
        public static string CleanupFiles { get; private set; } = "Deleting existing files and folders (excluding system folders)...";
        public static string CleanupComplete { get; private set; } = "Cleanup complete.";

        static Localization()
        {
            string lang = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName.ToLower();
            if (lang == "ru")
            {
                FormTitle = "USB-mgr";
                UsbDevices = "USB-устройства";
                IsoFile = "Файл ISO";
                Browse = "Обзор";
                Exit = "Выход";
                Start = "Старт";
                SelectUsbAndIso = "Пожалуйста, выберите USB-устройство и файл ISO.";
                EraseWarning = "Это удалит все данные на {0}. Продолжить?";
                Warning = "Предупреждение";
                NoPhysicalDrive = "Не удается найти физический диск.";
                StartingProcess = "Запуск процесса...";
                FormattingUsb = "Очистка USB-накопителя (удаление файлов без форматирования)...";
                MountingIso = "Монтирование ISO...";
                IsoMountedAt = "ISO смонтирован на {0}";
                CopyingFiles = "Копирование файлов на USB...";
                DismountingIso = "Размонтирование ISO...";
                BootableSuccess = "Загрузочный USB создан успешно!";
                Success = "Успех";
                ErrorFormat = "Ошибка: {0}";
                Error = "Ошибка";
                WaitingUsb = "Ожидание обнаружения USB-накопителя...";
                UsbDetectedAt = "USB-накопитель обнаружен на {0}";
                NoIsoLetter = "Не удается обнаружить букву смонтированного ISO-диска: {0}";
                NoUsbLetterAfterFormat = "Не удается найти букву USB-диска.";
                MountOutput = "Вывод монтирования: {0}";
                CreateFileFailed = "CreateFile не удался: {0}";
                DeviceIoControlFailed = "DeviceIoControl не удался: {0}";
                CleanupFiles = "Удаление существующих файлов и папок (кроме системных)...";
                CleanupComplete = "Очистка завершена.";
            }
            else if (lang == "ka")
            {
                FormTitle = "USB-mgr";
                UsbDevices = "USB მოწყობილობები";
                IsoFile = "ISO ფაილი";
                Browse = "დათვალიერება";
                Exit = "გასვლა";
                Start = "დაწყება";
                SelectUsbAndIso = "გთხოვთ, აირჩიოთ USB მოწყობილობა და ISO ფაილი.";
                EraseWarning = "ეს წაშლის ყველა მონაცემს {0}-ზე. გაგრძელება?";
                Warning = "გაფრთხილება";
                NoPhysicalDrive = "ფიზიკური დისკის პოვნა შეუძლებელია.";
                StartingProcess = "პროცესის დაწყება...";
                FormattingUsb = "USB დისკის გასუფთავება (მხოლოდ ფაილების წაშლა, ფორმატირების გარეშე)...";
                MountingIso = "ISO-ს მონტაჟი...";
                IsoMountedAt = "ISO მონტაჟირებულია {0}-ზე";
                CopyingFiles = "ფაილების კოპირება USB-ზე...";
                DismountingIso = "ISO-ს დემონტაჟი...";
                BootableSuccess = "ჩატვირთვადი USB წარმატებით შეიქმნა!";
                Success = "წარმატება";
                ErrorFormat = "შეცდომა: {0}";
                Error = "შეცდომა";
                WaitingUsb = "USB დრაივის აღმოჩენის ლოდინი...";
                UsbDetectedAt = "USB დრაივი აღმოჩენილია {0}-ზე";
                NoIsoLetter = "შეუძლებელია მონტაჟირებული ISO დისკის ასოს აღმოჩენა: {0}";
                NoUsbLetterAfterFormat = "USB დრაივის ასოს პოვნა შეუძლებელია.";
                MountOutput = "მონტაჟის გამოტანა: {0}";
                CreateFileFailed = "CreateFile წარუმატებელი: {0}";
                DeviceIoControlFailed = "DeviceIoControl წარუმატებელი: {0}";
                CleanupFiles = "არსებული ფაილებისა და საქაღალდეების წაშლა (სისტემურების გარდა)...";
                CleanupComplete = "გასუფთავება დასრულდა.";
            }
            else if (lang == "ab")
            {
                FormTitle = "USB-mgr";
                UsbDevices = "USB аҟәарақәа";
                IsoFile = "ISO афайл";
                Browse = "Аԥхьаара";
                Exit = "Аҭыҵра";
                Start = "Аԥхьаара";
                SelectUsbAndIso = "Ҳаҳәаԥштәи USB аҟәара ала аԥхьаара ала аISO афайл.";
                EraseWarning = "Ари аҭыҵра ала адыррақәа {0} ала. Аԥхьаара?";
                Warning = "Агәраҳәара";
                NoPhysicalDrive = "Афизикатә диск аԥхьаара аҭыҵра ала.";
                StartingProcess = "Апроцесс аԥхьаара...";
                FormattingUsb = "USB адиск агасупҭауеа (афайлқәа аҭыҵра, аформат ахалхҳәтәуп ала)...";
                MountingIso = "ISO амонтирование...";
                IsoMountedAt = "ISO амонтировано ала {0}";
                CopyingFiles = "Афайлқәа USB ала аԥхьარа...";
                DismountingIso = "ISO адисмонтирование...";
                BootableSuccess = "Аԥхьаара USB ала аԥхьაара аҭыҵра ала!";
                Success = "Аԥхьаара";
                ErrorFormat = "Агәра: {0}";
                Error = "Агәра";
                WaitingUsb = "USB адиск аԥхьაара ала аԥхьაара...";
                UsbDetectedAt = "USB адиск ала {0}";
                NoIsoLetter = "Амонтировано ISO адиск аԥхьაара ала: {0}";
                NoUsbLetterAfterFormat = "USB адиск аԥхьაара ала.";
                MountOutput = "Амонт аҭыҵра: {0}";
                CreateFileFailed = "CreateFile ала: {0}";
                DeviceIoControlFailed = "DeviceIoControl ала: {0}";
                CleanupFiles = "Адыррақәа аფайлқәа ала асақәа аҭыҵра (асистематә қәа ахалхҳәтәуп ала)...";
                CleanupComplete = "Агасупҭауеа аҭыҵра ала.";
            }
            else if (lang == "ace")
            {
                FormTitle = "USB-mgr";
                UsbDevices = "Alat USB";
                IsoFile = "Berkas ISO";
                Browse = "Pilih";
                Exit = "Keluar";
                Start = "Mulai";
                SelectUsbAndIso = "Silakan pilih perangkat USB dan berkas ISO.";
                EraseWarning = "Ini akan menghapus semua data pada {0}. Lanjutkan?";
                Warning = "Peringatan";
                NoPhysicalDrive = "Tidak dapat menemukan drive fisik.";
                StartingProcess = "Memulai proses...";
                FormattingUsb = "Membersihkan drive USB (hanya menghapus file, tanpa format)...";
                MountingIso = "Memasang ISO...";
                IsoMountedAt = "ISO dipasang di {0}";
                CopyingFiles = "Menyalin file ke USB...";
                DismountingIso = "Melepas ISO...";
                BootableSuccess = "USB bootable berhasil dibuat!";
                Success = "Sukses";
                ErrorFormat = "Kesalahan: {0}";
                Error = "Kesalahan";
                WaitingUsb = "Menunggu deteksi drive USB...";
                UsbDetectedAt = "Drive USB terdeteksi di {0}";
                NoIsoLetter = "Tidak dapat mendeteksi huruf drive ISO yang dipasang: {0}";
                NoUsbLetterAfterFormat = "Tidak dapat menemukan huruf drive USB.";
                MountOutput = "Output pemasangan: {0}";
                CreateFileFailed = "CreateFile gagal: {0}";
                DeviceIoControlFailed = "DeviceIoControl gagal: {0}";
                CleanupFiles = "Menghapus file dan folder yang ada (kecuali folder sistem)...";
                CleanupComplete = "Pembersihan selesai.";
            }
            // Default: English (already defined above)
        }
    }

    public partial class MainForm : Form
    {
        private ComboBox usbComboBox;
        private TextBox isoTextBox;
        private Button browseButton;
        private Button exitButton;
        private Button startButton;
        private TextBox logTextBox;
        private ProgressBar progressBar;
        private Label usbLabel;
        private Label isoLabel;
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
            this.usbLabel = new Label();
            this.isoLabel = new Label();

            // Form settings
            this.Text = Localization.FormTitle;
            this.Size = new Size(400, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;

            // USB Devices Label
            usbLabel.Text = Localization.UsbDevices;
            usbLabel.Location = new Point(10, 10);
            this.Controls.Add(usbLabel);

            // USB ComboBox
            usbComboBox.Location = new Point(10, 30);
            usbComboBox.Size = new Size(360, 20);
            usbComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            this.Controls.Add(usbComboBox);

            // ISO File Label
            isoLabel.Text = Localization.IsoFile;
            isoLabel.Location = new Point(10, 60);
            this.Controls.Add(isoLabel);

            // ISO TextBox
            isoTextBox.Location = new Point(10, 80);
            isoTextBox.Size = new Size(300, 20);
            this.Controls.Add(isoTextBox);

            // Browse Button
            browseButton.Text = Localization.Browse;
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
            exitButton.Text = Localization.Exit;
            exitButton.Location = new Point(10, 200);
            exitButton.Click += exitButton_Click;
            this.Controls.Add(exitButton);

            // Start Button
            startButton.Text = Localization.Start;
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
            this.Resize += (s, e) => { progressBar.Width = this.ClientSize.Width; };
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            RefreshUsbList();
        }

        private void RefreshUsbList()
        {
            usbComboBox.Items.Clear();
            foreach (var drive in DriveInfo.GetDrives())
            {
                if (drive.DriveType == DriveType.Removable && drive.IsReady)
                {
                    string item = $"{drive.Name} - {drive.VolumeLabel} ({Math.Round(drive.TotalSize / (1024.0 * 1024 * 1024), 1)} GB)";
                    usbComboBox.Items.Add(item);
                }
            }
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "ISO files|*.iso";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                isoTextBox.Text = ofd.FileName;
            }
        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private async void startButton_Click(object sender, EventArgs e)
        {
            if (isWorking) return;
            if (usbComboBox.SelectedIndex == -1 || string.IsNullOrEmpty(isoTextBox.Text))
            {
                MessageBox.Show(Localization.SelectUsbAndIso);
                return;
            }

            string selectedItem = usbComboBox.SelectedItem.ToString();
            string driveLetter = selectedItem.Substring(0, 2); // e.g., "D:"
            string isoPath = isoTextBox.Text;

            if (MessageBox.Show(string.Format(Localization.EraseWarning, driveLetter), Localization.Warning, MessageBoxButtons.YesNo) == DialogResult.No)
            {
                return;
            }

            string physicalDrive = GetPhysicalDriveFromLetter(driveLetter);
            if (physicalDrive == null)
            {
                MessageBox.Show(Localization.NoPhysicalDrive);
                return;
            }

            isWorking = true;
            startButton.Enabled = false;
            browseButton.Enabled = false;
            usbComboBox.Enabled = false;
            logTextBox.Clear();
            progressBar.Value = 0;

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
                AppendLog(Localization.StartingProcess);

                // Step 1: Get USB drive letter & clean files (NO formatting, NO FAT32 check!)
                string usbLetter = GetDriveLetterFromPhysical(physicalDrive);
                if (string.IsNullOrEmpty(usbLetter))
                {
                    throw new Exception(Localization.NoUsbLetterAfterFormat);
                }

                AppendLog($"USB drive letter: {usbLetter}");
                UpdateProgress(10);

                // ✅ ფაილური სისტემის შემოწმება სრულიად გამორთულია — გრძელდება მხოლოდ წაშლა

                // Delete all non-system files & folders
                AppendLog(Localization.CleanupFiles);
                try
                {
                    var root = new DirectoryInfo(usbLetter);
                    foreach (var file in root.GetFiles("*", SearchOption.TopDirectoryOnly))
                    {
                        if ((file.Attributes & (FileAttributes.Hidden | FileAttributes.System)) == 0)
                        {
                            file.Delete();
                        }
                    }
                    foreach (var dir in root.GetDirectories("*", SearchOption.TopDirectoryOnly))
                    {
                        if ((dir.Attributes & (FileAttributes.Hidden | FileAttributes.System)) == 0)
                        {
                            dir.Delete(true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"⚠️ Warning during cleanup: {ex.Message}");
                }

                AppendLog(Localization.CleanupComplete);
                UpdateProgress(30);

                // Step 2: Mount ISO
                AppendLog(Localization.MountingIso);
                string mountCommand = $"-Command \"$di = Mount-DiskImage -ImagePath '{isoPath.Replace("\\", "\\\\")}' -PassThru; (Get-Volume -DiskImage $di).DriveLetter\"";
                string mountOutput = await RunProcessAsync("powershell.exe", mountCommand, false);
                AppendLog(string.Format(Localization.MountOutput, mountOutput));
                string mountedLetterStr = mountOutput.Trim();
                if (!string.IsNullOrEmpty(mountedLetterStr) &&
                    mountedLetterStr.IndexOf("Error", StringComparison.OrdinalIgnoreCase) == -1)
                {
                    string mountedLetter = mountedLetterStr + ":\\";
                    AppendLog(string.Format(Localization.IsoMountedAt, mountedLetter));
                    UpdateProgress(40);

                    // Step 3: Copy files (progress-aware)
                    AppendLog(Localization.CopyingFiles);
                    CopyDirectoryWithProgress(mountedLetter, usbLetter + "\\");
                    UpdateProgress(80);

                    // Step 4: Dismount ISO
                    AppendLog(Localization.DismountingIso);
                    string dismountOutput = await RunProcessAsync("powershell.exe", $"-Command Dismount-DiskImage -ImagePath '{isoPath.Replace("\\", "\\\\")}'", false);
                    AppendLog(dismountOutput);
                    UpdateProgress(100);

                    AppendLog(Localization.BootableSuccess);
                    InvokeMessageBox(Localization.Success, Localization.BootableSuccess);
                }
                else
                {
                    throw new Exception(string.Format(Localization.NoIsoLetter, mountOutput));
                }
            }
            catch (OperationCanceledException)
            {
                AppendLog("Operation canceled by user.");
                UpdateProgress(0);
            }
            catch (Exception ex)
            {
                AppendLog(string.Format(Localization.ErrorFormat, ex.Message));
                UpdateProgress(0);
                InvokeMessageBox(Localization.Error, string.Format(Localization.ErrorFormat, ex.Message));
            }
        }

        private void CopyDirectoryWithProgress(string sourceDir, string targetDir)
        {
            var files = Directory.GetFiles(sourceDir, "*.*", SearchOption.AllDirectories);
            int totalFiles = files.Length;
            int copied = 0;

            foreach (string file in files)
            {
                string relativePath = file.Substring(sourceDir.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string destFile = Path.Combine(targetDir, relativePath);
                string destDir = Path.GetDirectoryName(destFile);
                if (!string.IsNullOrEmpty(destDir) && !Directory.Exists(destDir))
                {
                    Directory.CreateDirectory(destDir);
                }

                try
                {
                    File.Copy(file, destFile, true);
                }
                catch (Exception ex)
                {
                    AppendLog($"⚠️ Skip (copy failed): {relativePath} → {ex.Message}");
                }

                copied++;
                int progress = 40 + (int)(40 * (copied / (double)totalFiles)); // From 40% to 80%
                UpdateProgress(progress);
            }
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
                output = stdOut.Trim() + (string.IsNullOrEmpty(stdErr) ? "" : "\nError: " + stdErr.Trim());
            }
            await Task.Run(() => process.WaitForExit());
            return output;
        }

        private void AppendLog(string message)
        {
            this.Invoke(new Action(() =>
            {
                logTextBox.AppendText(DateTime.Now.ToString("HH:mm:ss") + " | " + message + "\r\n");
                logTextBox.ScrollToCaret();
            }));
        }

        private void UpdateProgress(int value)
        {
            this.Invoke(new Action(() => progressBar.Value = Math.Min(value, 100)));
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
                    AppendLog(string.Format(Localization.CreateFileFailed, Marshal.GetLastWin32Error()));
                    return -1;
                }

                STORAGE_DEVICE_NUMBER sdn;
                uint bytesReturned;
                if (DeviceIoControl(handle, IOCTL_STORAGE_GET_DEVICE_NUMBER, IntPtr.Zero, 0, out sdn, Marshal.SizeOf(typeof(STORAGE_DEVICE_NUMBER)), out bytesReturned, IntPtr.Zero))
                {
                    return sdn.DeviceNumber;
                }
                else
                {
                    AppendLog(string.Format(Localization.DeviceIoControlFailed, Marshal.GetLastWin32Error()));
                    return -1;
                }
            }
        }

        private string GetPhysicalDriveFromLetter(string letter)
        {
            int diskNum = GetDiskNumberFromLetter(letter);
            if (diskNum == -1)
            {
                return null;
            }
            return @"\\.\PhysicalDrive" + diskNum;
        }

        private string GetDriveLetterFromPhysical(string physical)
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
            Thread.CurrentThread.CurrentCulture = CultureInfo.CurrentCulture;
            Thread.CurrentThread.CurrentUICulture = CultureInfo.CurrentUICulture;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}