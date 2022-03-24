using ExcelDataReader;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
namespace GonderBayrakExcel
{
    public partial class Form1 : Form
    {
        DataTableCollection dtc;
        List<string> createdFolders = new List<string>();
        DateTime fileAndFolderCreateDate = new DateTime();
        DateTime fileAndFolderEditDate = new DateTime();
        DateTime processStartDate = new DateTime();
        public class ExcelRow
        {
            public int Id { get; set; }

            public string Po { get; set; }
            public string Liste { get; set; }
            public string CustCode { get; set; }
            public string Adi { get; set; }
            public string TabelaYazisi { get; set; }
            public string Ilce { get; set; }
            public string Il { get; set; }
            public string BayiCode { get; set; }
            public string BayiAdi { get; set; }
            public string BayiUnvan { get; set; }
            public string BayiAdres { get; set; }
            public int FolderId { get; set; }


        }

        List<ExcelRow> excelList = new List<ExcelRow>();

        BackgroundWorker backgroundWorker = new BackgroundWorker();
        public Form1()
        {
            InitializeComponent();
            label1.Visible = false;
            button1.Visible = true;
            progressBar1.Visible = false;
            processStartDate = new DateTime(2020, 07, 10, 11, 00, 00);
            fileAndFolderCreateDate = processStartDate;
            fileAndFolderEditDate = processStartDate;

        }


        private DateTime GetRandomDate(DateTime date)
        {
            Random rnd = new Random();

            double numara = rnd.Next(5, 15);
            DateTime resultDate = date.AddMinutes(numara);

            return resultDate;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openfile = new OpenFileDialog() { Filter = "Excel Dosyaları |*.xlsx|Excel Dosyaları 97-2003|*.xls", Title = "EXCEL DOSYALARI" })
                {

                    if (openfile.ShowDialog() == DialogResult.OK)
                    {
                        label1.Text = label1.Text + Path.GetFileName(openfile.FileName);
                        label1.Visible = true;
                        button1.Visible = false;


                        progressBar1.Visible = true;
                        // Set Minimum to 1 to represent the first file being copied.
                        progressBar1.Minimum = 1;

                        // Set the initial value of the ProgressBar.
                        progressBar1.Value = 1;
                        // Set the Step property to a value of 1 to represent each file being copied.
                        progressBar1.Step = 1;

                        var excelfileName = openfile.FileName; //"../../MuratKlasor/Yelken Bayrak-Gönder-Bayrak.xlsx";

                        DataTableCollection dtc = GetExcelFileCompenent(excelfileName);
                        foreach (DataTable dt in dtc)
                        {

                            int rowId = 1;
                            int folderId = 0;

                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                //if (i == 30)
                                //{
                                //    var a = 2 + 1;

                                //}

                                if (excelList.Where(x => x.BayiAdi == dt.Rows[i]["BAYİ/DİST ADI"]).Count() == 0)
                                {
                                    folderId++;
                                }
                                else
                                {
                                    folderId = excelList.Where(x => x.BayiAdi == dt.Rows[i]["BAYİ/DİST ADI"]).First().FolderId;
                                }
                                ExcelRow excelRow = new ExcelRow();
                                excelRow.Id = rowId++;
                                excelRow.Po = dt.Rows[i]["PO"].ToString();
                                excelRow.Liste = dt.Rows[i]["LİSTE"].ToString();
                                excelRow.CustCode = dt.Rows[i]["Cust Code"].ToString();
                                excelRow.Adi = dt.Rows[i]["Adı"].ToString().Replace("/", "-").Replace(":", "-");
                                excelRow.TabelaYazisi = dt.Rows[i]["Tabelada Yazılması İstenilen İsim "].ToString();
                                excelRow.Ilce = dt.Rows[i]["İlçe"].ToString();
                                excelRow.Il = dt.Rows[i]["İl"].ToString();
                                excelRow.BayiCode = dt.Rows[i]["BAYİ/DİST KODU"].ToString();
                                excelRow.BayiAdi = dt.Rows[i]["BAYİ/DİST ADI"].ToString();
                                excelRow.BayiUnvan = dt.Rows[i]["BAYİ/DİST UNVAN"].ToString();
                                excelRow.BayiAdres = dt.Rows[i]["BAYİ/DİST ADRES"].ToString();
                                excelRow.FolderId = folderId;

                                if(excelRow.Adi != "")
                                {
                                    excelList.Add(excelRow);
                                }
                                folderId = excelList.OrderByDescending(x => x.FolderId).First().FolderId;

                            }
                            //comboBox1.Items.Add(table.TableName);
                        }

                        // Set Maximum to the total number of files to copy.
                        progressBar1.Maximum = excelList.Count;

                        foreach (var excelRow in excelList)
                        {
                            //İptal butonunu becerebilirsem. İptal olunca oluşturulan klasörleri tümüyle silsin diye eklendi burası
                            //if (cancelProcess)
                            //{

                            //    foreach (var createdFolder in createdFolders)
                            //    {
                            //        if (Directory.Exists(createdFolder))
                            //        {
                            //            Directory.Delete(createdFolder);
                            //        }
                            //    }
                            //    break;

                            //}

                            if (excelRow.Adi != "")
                            {
                                DateTime folderCreateDate = DateTime.Now;
                                var folderPath = $"../../IslenmisKlasor/({excelRow.FolderId})-{excelRow.BayiAdi}";
                                if (!Directory.Exists(folderPath))
                                {
                                    fileAndFolderCreateDate = GetRandomDate(fileAndFolderCreateDate);
                                    folderCreateDate = fileAndFolderCreateDate;

                                    Directory.CreateDirectory(folderPath);
                                    createdFolders.Add(folderPath);
                                }

                                var createdFilePath = CreatePdf(excelRow, folderPath);

                                if (File.Exists(createdFilePath))
                                {
                                    fileAndFolderCreateDate = GetRandomDate(fileAndFolderCreateDate);
                                    fileAndFolderEditDate = GetRandomDate(fileAndFolderEditDate);

                                    FileInfo fileInfo = new FileInfo(@"" + createdFilePath);

                                    fileInfo.CreationTime = fileAndFolderCreateDate;
                                    fileInfo.LastWriteTime = fileAndFolderEditDate;
                                    fileInfo.LastAccessTime = fileAndFolderEditDate;
                                    Directory.SetLastWriteTime(folderPath, fileAndFolderEditDate);

                                }


                                if (Directory.Exists(folderPath))
                                {
                                    
                                    Directory.SetCreationTime(folderPath, folderCreateDate);
                                    //Directory.SetLastAccessTime(folderPath, folderCreateDate);
                                }
                            }

                            progressBar1.PerformStep();
                        }


                    }
                }

            }
            catch (Exception ex)
            {

                throw ex;
            }
        }


        private DataTableCollection GetExcelFileCompenent(string excelFilePath)
        {

            using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader excelreader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = excelreader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (x) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    }

                    );
                    return dtc = result.Tables;

                }

            }

        }

        public static string CreatePdf(ExcelRow excelRow, string folderPath)
        {
            string path = "";
            try
            {
                iTextSharp.text.Document doc = new iTextSharp.text.Document(PageSize.A4.Rotate(), 0, 0, 10, 0);
                var randomNumber = new Random().Next();


                path = folderPath + $"/({excelRow.Id})-{excelRow.TabelaYazisi.Replace("/", "-").Replace(":", "-")}.pdf";

                //BaseFont turkishCharacterFontSettings = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, "windows-1254", BaseFont.EMBEDDED);
                //BaseFont turkishCharacterFontSettings = BaseFont.CreateFont(BaseFont.HELVETICA, "windows-1254", BaseFont.EMBEDDED);
                BaseFont turkishCharacterFontSettings = BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, "windows-1254", BaseFont.EMBEDDED);

                iTextSharp.text.Font textNormalFont20B = new iTextSharp.text.Font(turkishCharacterFontSettings, 22f, iTextSharp.text.Font.BOLD);
                iTextSharp.text.Font textNormalFont20 = new iTextSharp.text.Font(turkishCharacterFontSettings, 20f, iTextSharp.text.Font.BOLD);

                var pdfWriter = PdfWriter.GetInstance(doc, new FileStream(path, FileMode.Create));
                doc.Open();

                PdfPTable tableForLogo = new PdfPTable(1) { WidthPercentage = 100 };
                var logo = iTextSharp.text.Image.GetInstance("../../img/gonderBayrakBilgiler.png");
                logo.ScaleToFit(800f, 800f);
                tableForLogo.AddCell(new PdfPCell(logo) { Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
                //tableForLogo.AddCell(new PdfPCell(logo) { Padding = 10, Border = 0, HorizontalAlignment = Element.ALIGN_CENTER });
                doc.Add(tableForLogo);

                PdfPTable textOfFile = new PdfPTable(5) { WidthPercentage = 100 };
                textOfFile.AddCell(new PdfPCell(new Phrase("  ", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase("  ", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"PO:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.Po}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Liste:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.Liste}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Cust Code:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.CustCode}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Adı:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.Adi}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Tabela Yazısı:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.TabelaYazisi}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("İl-İlçe:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.Il}-{excelRow.Ilce}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Bayi/Dist Kodu:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.BayiCode}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Bayi/Dist Adı:", textNormalFont20B)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.BayiAdi}", textNormalFont20B)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Bayi/Dist Unvan:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.BayiUnvan}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                textOfFile.AddCell(new PdfPCell(new Phrase("Bayi/Dist Adres:", textNormalFont20)) { BorderWidth = 0 });
                textOfFile.AddCell(new PdfPCell(new Phrase($"{excelRow.BayiAdres}", textNormalFont20)) { BorderWidth = 0, Colspan = 4 });
                doc.Add(textOfFile);


                doc.Close();


                return path;
                //BREAK POİNT NOKTASI
            }
            catch (Exception ex)
            {

                if (File.Exists(path))
                    File.Delete(path);

                throw ex;
            }
        }


        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }

    }
}
