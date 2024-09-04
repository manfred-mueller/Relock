using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ZXing;
using Version = System.Version;

namespace Relock
{
    internal static class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(int dwProcessId);
        private const int ATTACH_PARENT_PROCESS = -1;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool FreeConsole();

        [STAThread]
        private static void Main(string[] args)
        {
            // Try to attach to the parent process's console, if it exists
            bool isConsoleAttached = AttachConsole(ATTACH_PARENT_PROCESS);

            if (args.Length > 0)
            {
                string drive = args[0].ToLower();
                switch (drive)
                {
                    case "/register":
                        RegisterInRegistry();
                        break;
                    case "/unregister":
                        UnregisterFromRegistry();
                        break;
                    default:
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);
                        getRecoveryKey(drive);
                        lockDrive(drive);
                        break;
                }
            }
            else
            {
                if (isConsoleAttached)
                {
                    Version version = Assembly.GetExecutingAssembly().GetName().Version;

                    string relockVersion = $"{version.Major}.{version.Minor}.{version.Build}";
                    Console.WriteLine("\n" +
                        String.Format(Properties.Resources.Relock0ReLockABitlockerEnabledDrive2024ManfredMueller, relockVersion) + "\n\n" +
                        Properties.Resources.UsageRelockRegisterUnregister + "\n" +
                        "/register\t" + Properties.Resources.RegisterInTheExplorerContextMenu + "\n" +
                        "/unregister\t" + Properties.Resources.UnregisterFromTheExplorerContextMenu
                    );
                    FreeConsole();
                    Environment.Exit(0);
                }
            }
        }

        public static void RegisterInRegistry()
        {
            string keyName = "HKEY_LOCAL_MACHINE\\SOFTWARE\\Classes\\Drive\\shell\\relock-bde";
            string _vDefault = Properties.Resources.RelockThisDrive;
            const string _vAppliesTo = "System.Volume.BitLockerProtection:=System.Volume.BitLockerProtection#On OR System.Volume.BitLockerProtection:=System.Volume.BitLockerProtection#Encrypting OR System.Volume.BitLockerProtection:=System.Volume.BitLockerProtection#Suspended";
            const string _vMultiSelectModel = "Single";
            string appPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string appIcon = appPath + ",0";  // Assuming the icon is the first resource in the executable

            try
            {
                Registry.SetValue(keyName, "", _vDefault);
                Registry.SetValue(keyName, "AppliesTo", _vAppliesTo);
                Registry.SetValue(keyName, "MultiSelectModel", _vMultiSelectModel);
                Registry.SetValue(keyName, "Icon", appIcon);

                string commandKeyName = keyName + "\\command";
                string _vCommandValue = appPath + " %1";

                Registry.SetValue(commandKeyName, "", _vCommandValue);

                MessageBox.Show(Properties.Resources.ProgramSuccessfullyRegisteredInTheRegistry, "Relock", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.ErrorWhileUpdatingRegistry + ex.Message, "Relock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void UnregisterFromRegistry()
        {
            try
            {
                RegistryKey key = Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes\\Drive\\shell", true);
                if (key != null)
                {
                    key.DeleteSubKeyTree("relock-bde", false);
                    key.Close();

                    MessageBox.Show(Properties.Resources.ProgramSuccessfullyUnregisteredFromTheRegistry, "Relock", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(Properties.Resources.RegistryKeyNotFound, "Relock", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Properties.Resources.ErrorWhileUpdatingRegistry + ex.Message, "Relock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void getRecoveryKey(string driveLetter)
        {
            try
            {
                driveLetter = driveLetter.Replace("\\", "");

                var psi = new ProcessStartInfo("manage-bde", string.Format("-protectors -get {0} -Type recoverypassword", driveLetter))
                {
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                };

                StringBuilder outputBuilder = new StringBuilder();

                using (Process process = new Process())
                {
                    process.StartInfo = psi;
                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            outputBuilder.AppendLine(e.Data);
                        }
                    };

                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            outputBuilder.AppendLine(Relock.Properties.Resources.OutputError + e.Data);
                        }
                    };

                    process.Start();
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    process.WaitForExit();
                }

                string output = outputBuilder.ToString();

                // Filter lines that end with a number and do not contain "ersion"
                string[] lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string line in lines)
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(line, @"[0-9]$") && !line.Contains("ersion"))
                    {
                        // Remove spaces and get the result
                        string result = line.Replace(" ", string.Empty);

                        // Generate QR code image
                        string qrData = $"{result}";
                        var qrWriter = new BarcodeWriter
                        {
                            Format = BarcodeFormat.QR_CODE,
                            Options = new ZXing.Common.EncodingOptions
                            {
                                Width = 400,
                                Height = 400,
                                Margin = 1
                            }
                        };

                        using (var qrImage = qrWriter.Write(qrData))
                        {
                            // Convert the icon to a bitmap
                            using (var iconBitmap = Properties.Resources.relock.ToBitmap())
                            {
                                using (Graphics graphics = Graphics.FromImage(qrImage))
                                {
                                    // Draw the icon on the QR code, centered on the white square
                                    int iconX = (qrImage.Width - iconBitmap.Width) / 2;
                                    int iconY = (qrImage.Height - iconBitmap.Height) / 2;
                                    graphics.DrawImage(iconBitmap, new Point(iconX, iconY));
                                }

                                // Create the PictureBox to display the combined image
                                using (var pictureBox = new PictureBox())
                                {
                                    pictureBox.Image = qrImage;
                                    pictureBox.SizeMode = PictureBoxSizeMode.AutoSize;
                                    pictureBox.Refresh(); // Force refresh to ensure the image is displayed

                                    // Create the label for the recovery key
                                    Label recoveryKeyLabel = new Label
                                    {
                                        AutoSize = false,
                                        Width = pictureBox.Width,
                                        TextAlign = ContentAlignment.MiddleCenter,
                                        MaximumSize = new Size(pictureBox.Width, 0), // Ensure the label does not exceed the width of the PictureBox
                                        Height = 50, // Ensure enough height to fit the text
                                        Text = result
                                    };

                                    // Adjust font size to fit the label width
                                    System.Drawing.Font font = AdjustFontToFitLabel(recoveryKeyLabel, result);
                                    recoveryKeyLabel.Font = font;

                                    // Create the label for the drive letter
                                    Label driveLabel = new Label
                                    {
                                        AutoSize = false,
                                        Width = pictureBox.Width,
                                        TextAlign = ContentAlignment.MiddleCenter,
                                        MaximumSize = new Size(pictureBox.Width, 0),
                                        Height = 30, // Height of the label above the QR code
                                        Text = string.Format(Properties.Resources.RecoveryCodeForDrive0, driveLetter.ToUpper())
                                    };

                                    // Adjust font size to fit the label width
                                    System.Drawing.Font driveFont = AdjustFontToFitLabel(driveLabel, driveLabel.Text);
                                    driveLabel.Font = driveFont;

                                    // Create the form
                                    Form form = new Form
                                    {
                                        Icon = Properties.Resources.relock,
                                        Text = Properties.Resources.RecoveryKey,
                                        AutoSize = true,
                                        AutoSizeMode = AutoSizeMode.GrowAndShrink,
                                        StartPosition = FormStartPosition.CenterScreen
                                    };

                                    // Add controls to the form in the correct order
                                    form.Controls.Add(driveLabel);
                                    form.Controls.Add(pictureBox);
                                    form.Controls.Add(recoveryKeyLabel);

                                    // Create and configure the PrintButton
                                    Button printButton = new Button
                                    {
                                        Text = Properties.Resources.Print,
                                        AutoSize = true
                                    };

                                    printButton.Click += (sender, e) =>
                                    {
                                        PrintDocument printDocument = new PrintDocument();
                                        printDocument.PrintPage += (s, ev) =>
                                        {
                                            float margin = 10;
                                            float xCenter = (ev.PageBounds.Width - pictureBox.Width) / 2;
                                            float y = margin;

                                            ev.Graphics.DrawString(driveLabel.Text, driveLabel.Font, Brushes.Black, xCenter, y);
                                            y += driveLabel.Height + margin;
                                            ev.Graphics.DrawImage(pictureBox.Image, xCenter, y);
                                            y += pictureBox.Height + margin;
                                            ev.Graphics.DrawString(recoveryKeyLabel.Text, recoveryKeyLabel.Font, Brushes.Black, xCenter, y);
                                        };

                                        using (PrintDialog printDialog = new PrintDialog())
                                        {
                                            printDialog.Document = printDocument;
                                            if (printDialog.ShowDialog() == DialogResult.OK)
                                            {
                                                printDocument.Print();
                                            }
                                        }
                                    };

                                    // Create and configure the SaveButton
                                    Button saveButton = new Button
                                    {
                                        Text = Properties.Resources.Save,
                                        AutoSize = true
                                    };

                                    saveButton.Click += (sender, e) =>
                                    {
                                        SaveFileDialog saveFileDialog = new SaveFileDialog
                                        {
                                            CreatePrompt = true,
                                            OverwritePrompt = true,
                                            Filter = Relock.Properties.Resources.PDFFilesPdfPdf,
                                            DefaultExt = "pdf",
                                            FileName = "RecoveryKey.pdf"
                                        };

                                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                                        {
                                            // Create the PDF document
                                            using (var pdfDocument = new iTextSharp.text.Document())
                                            {
                                                PdfWriter.GetInstance(pdfDocument, new FileStream(saveFileDialog.FileName, FileMode.Create));
                                                pdfDocument.Open();

                                                // Add the drive label text
                                                var driveLabelParagraph = new iTextSharp.text.Paragraph(driveLabel.Text, FontFactory.GetFont("Arial", driveLabel.Font.Size, iTextSharp.text.Font.BOLD))
                                                {
                                                    Alignment = Element.ALIGN_CENTER,
                                                    SpacingAfter = 20f
                                                };
                                                pdfDocument.Add(driveLabelParagraph);

                                                // Add the QR code image
                                                using (var qrStream = new MemoryStream())
                                                {
                                                    qrImage.Save(qrStream, System.Drawing.Imaging.ImageFormat.Png);
                                                    var qrPdfImage = iTextSharp.text.Image.GetInstance(qrStream.ToArray());
                                                    qrPdfImage.Alignment = Element.ALIGN_CENTER;
                                                    pdfDocument.Add(qrPdfImage);
                                                }

                                                // Add the recovery key text
                                                var recoveryKeyParagraph = new iTextSharp.text.Paragraph(recoveryKeyLabel.Text, FontFactory.GetFont("Arial", recoveryKeyLabel.Font.Size, iTextSharp.text.Font.BOLD))
                                                {
                                                    Alignment = Element.ALIGN_CENTER,
                                                    SpacingBefore = 20f
                                                };
                                                pdfDocument.Add(recoveryKeyParagraph);

                                                pdfDocument.Close();
                                            }
                                        }
                                    };

                                    form.Controls.Add(printButton);
                                    form.Controls.Add(saveButton);

                                    // Adjust layout
                                    driveLabel.Location = new Point(0, 10);
                                    pictureBox.Location = new Point(0, driveLabel.Bottom + 10);
                                    recoveryKeyLabel.Location = new Point(0, pictureBox.Bottom + 10);
                                    saveButton.Location = new Point(form.ClientSize.Width / 2 - saveButton.Width - 5, recoveryKeyLabel.Bottom + 10);
                                    printButton.Location = new Point(form.ClientSize.Width / 2 + 5, recoveryKeyLabel.Bottom + 10);

                                    form.ShowDialog();
                                }
                            }
                        }

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMessage = string.Format(Properties.Resources.ErrorRetrievingKey, ex.Message);
                MessageBox.Show(errorMessage, Properties.Resources.Error, MessageBoxButtons.OK, MessageBoxIcon.Error);

                // Copy the error message to the clipboard
                Clipboard.SetText(errorMessage);
            }
        }

        private static System.Drawing.Font AdjustFontToFitLabel(Label label, string text)
        {
            // Define the maximum font size to try
            float fontSize = 20; // Initial font size
            System.Drawing.Font font = new System.Drawing.Font("Arial", fontSize, FontStyle.Bold);
            SizeF textSize;

            // Measure text size and adjust font size
            using (Graphics g = label.CreateGraphics())
            {
                do
                {
                    fontSize--;
                    font = new System.Drawing.Font("Arial", fontSize, FontStyle.Bold);
                    textSize = g.MeasureString(text, font);
                }
                while (textSize.Width > label.Width && fontSize > 1);
            }

            return font;
        }
        public static void lockDrive(string driveLetter)
        {
            try
            {
                driveLetter = driveLetter.Replace("\\", "");

                var psi = new ProcessStartInfo("manage-bde", string.Format("-lock {0} -ForceDismount", driveLetter))
                { CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden };
                Process.Start(psi);
            }
            catch (Exception exc)
            {
                MessageBox.Show(string.Format( Properties.Resources.FailedToLockDriveRN0, exc.Message), "Relock",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
