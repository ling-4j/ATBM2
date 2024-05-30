using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace ATBM2
{
    public partial class MainWindow : Window
    {
        private string originalText;
        private byte[] originalSignature;
        private DSACryptoServiceProvider dsa;

        public MainWindow()
        {
            InitializeComponent();
            dsa = new DSACryptoServiceProvider();
        }

        private void LoadWordDocument(string filePath, RichTextBox richTextBox)
        {
            string rtfContent = ConvertWordToRtf(filePath);
            richTextBox.Document = ConvertRtfToFlowDocument(rtfContent);
        }

        private string ConvertWordToRtf(string filePath)
        {
            var wordApp = new Word.Application();
            var wordDoc = wordApp.Documents.Open(filePath);

            string tempFile = Path.GetTempFileName();
            wordDoc.SaveAs2(tempFile, Word.WdSaveFormat.wdFormatRTF);

            wordDoc.Close();
            wordApp.Quit();

            string rtfContent;
            using (var streamReader = new StreamReader(tempFile))
            {
                rtfContent = streamReader.ReadToEnd();
            }

            File.Delete(tempFile);
            return rtfContent;
        }

        private FlowDocument ConvertRtfToFlowDocument(string rtfContent)
        {
            var flowDocument = new FlowDocument();
            var textRange = new TextRange(flowDocument.ContentStart, flowDocument.ContentEnd);

            using (var memoryStream = new MemoryStream(Encoding.UTF8.GetBytes(rtfContent)))
            {
                textRange.Load(memoryStream, DataFormats.Rtf);
            }

            return flowDocument;
        }

        private string GetTextFromRichTextBox(RichTextBox richTextBox)
        {
            return new TextRange(richTextBox.Document.ContentStart, richTextBox.Document.ContentEnd).Text.Trim();
        }

        private byte[] ComputeHash(string input)
        {
            using (SHA1 sha1 = SHA1.Create())
            {
                return sha1.ComputeHash(Encoding.UTF8.GetBytes(input));
            }
        }

        private void btn_file_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents|*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                LoadWordDocument(filePath, rtbVBK);
            }
        }

        private void btn_fileVB_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Word Documents|*.docx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string filePath = openFileDialog.FileName;
                LoadWordDocument(filePath, rtbChkVBK);
            }
        }

        private void btn_fileCK_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                rtbChkCK.Document.Blocks.Clear();
                rtbChkCK.AppendText(File.ReadAllText(openFileDialog.FileName));
            }
        }

        private void btn_sign_Click(object sender, RoutedEventArgs e)
        {
            originalText = GetTextFromRichTextBox(rtbVBK);
            byte[] hash = ComputeHash(originalText);
            originalSignature = dsa.SignHash(hash, CryptoConfig.MapNameToOID("SHA1"));

            rtbDSA.Document.Blocks.Clear();
            rtbDSA.AppendText(BitConverter.ToString(hash).Replace("-", ""));
            rtbCK.Document.Blocks.Clear();
            rtbCK.AppendText(BytesToHex(originalSignature));
        }

        private void btn_chkSign_Click(object sender, RoutedEventArgs e)
        {
            if (originalSignature != null && rtbChkCK.Document.Blocks.Count > 0)
            {
                string chkText = GetTextFromRichTextBox(rtbChkVBK);
                byte[] chkHash = ComputeHash(chkText);
                rtbChkDSA.Document.Blocks.Clear();
                rtbChkDSA.AppendText(BytesToHex(chkHash));

                string originalSignatureText = GetTextFromRichTextBox(rtbChkCK);
                byte[] originalSignatureBytes = HexToBytes(originalSignatureText);

                if (dsa.VerifyHash(chkHash, CryptoConfig.MapNameToOID("SHA1"), originalSignatureBytes))
                {
                    if (originalText == chkText)
                    {
                        rtbMess.Document.Blocks.Clear();
                        rtbMess.AppendText("Chữ ký đúng; văn bản không thay đổi");
                    }
                    else
                    {
                        rtbMess.Document.Blocks.Clear();
                        rtbMess.AppendText("Chữ ký đúng; văn bản đã thay đổi");
                    }
                }
                else
                {
                    rtbMess.Document.Blocks.Clear();
                    rtbMess.AppendText("Chữ ký sai");
                }
            }
            else
            {
                MessageBox.Show("Vui lòng chọn file văn bản và ký trước khi kiểm tra chữ ký.", "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btn_move_Click(object sender, RoutedEventArgs e)
        {
            rtbChkVBK.Document.Blocks.Clear();
            rtbChkCK.Document.Blocks.Clear();

            MemoryStream streamVBK = new MemoryStream();
            TextRange sourceRangeVBK = new TextRange(rtbVBK.Document.ContentStart, rtbVBK.Document.ContentEnd);
            sourceRangeVBK.Save(streamVBK, DataFormats.Rtf);
            TextRange destinationRangeVBK = new TextRange(rtbChkVBK.Document.ContentStart, rtbChkVBK.Document.ContentEnd);
            destinationRangeVBK.Load(streamVBK, DataFormats.Rtf);
            streamVBK.Close();

            MemoryStream streamCK = new MemoryStream();
            TextRange sourceRangeCK = new TextRange(rtbCK.Document.ContentStart, rtbCK.Document.ContentEnd);
            sourceRangeCK.Save(streamCK, DataFormats.Rtf);
            TextRange destinationRangeCK = new TextRange(rtbChkCK.Document.ContentStart, rtbChkCK.Document.ContentEnd);
            destinationRangeCK.Load(streamCK, DataFormats.Rtf);
            streamCK.Close();
        }

        private void btn_save_Click(object sender, RoutedEventArgs e)
        {
            if (originalSignature != null)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    try
                    {
                        string signatureText = BytesToHex(originalSignature);
                        File.WriteAllText(filePath, signatureText);
                        MessageBox.Show("Chữ ký đã được lưu thành công.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Đã xảy ra lỗi khi lưu chữ ký: " + ex.Message, "Lỗi", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            else
            {
                MessageBox.Show("Chưa có chữ ký để lưu.", "Thông báo", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private byte[] HexToBytes(string hex)
        {
            if (hex.Length % 2 != 0)
                throw new ArgumentException("Invalid hex string");

            byte[] bytes = new byte[hex.Length / 2];
            for (int i = 0; i < hex.Length; i += 2)
            {
                bytes[i / 2] = Convert.ToByte(hex.Substring(i, 2), 16);
            }
            return bytes;
        }

        private string BytesToHex(byte[] bytes)
        {
            StringBuilder sb = new StringBuilder(bytes.Length * 2);
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("X2"));
            }
            return sb.ToString();
        }
    }
}

