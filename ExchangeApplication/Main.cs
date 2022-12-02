using Microsoft.Web.WebView2.Core;
using System;
using System.Globalization;
using ClosedXML.Excel;
using System.Reflection.Metadata;
using System.Windows.Forms;

namespace ExchangeApplication
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            webView21.Source = new Uri("https://www.tcmb.gov.tr/kurlar/kurlar_tr.html");

            dataGridView1.ColumnCount = 7;
            dataGridView1.Columns[0].Name = "Doviz Kodu"; // dataGridViev sütun açıklamaları tanımlanması
            dataGridView1.Columns[1].Name = "Birim";
            dataGridView1.Columns[2].Name = "Doviz Cinsi";
            dataGridView1.Columns[3].Name = "Doviz Alis";
            dataGridView1.Columns[4].Name = "Doviz Satis";
            dataGridView1.Columns[5].Name = "Elektif Alis";
            dataGridView1.Columns[6].Name = "Elektif Satis";
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            export.Enabled = false;     //excel dışa aktarma butonu herhangi bir data yüklenmediğinde basılamaz durumda olmalı

        }

        private async void submit_Click(object sender, EventArgs e)
        {
            submit.Enabled = false; //bir data yüklenirken ikinci bir istekte bulunulmaması amacıyla süreç bitene kadar buton basılamaz durumda olacaktır
            string buildedRequestXmlUrl = dateTimePicker1.Value.ToString("yyyyMM") + '/' + dateTimePicker1.Value.ToString("ddMMyyyy") + ".xml";
            string buildedRequestXslUrl = dateTimePicker1.Value.ToString("yyyyMM") + '/' + "isokur-new" + ".xsl";
            _ = webView21.CoreWebView2.ExecuteScriptAsync(
                "displayResult(\"" + buildedRequestXslUrl + "\",\"" + buildedRequestXmlUrl + "\"); " +
                "$(\"#back\").show();" +
                "$(\"#data\").show();" +
                "$(\"#calendar_report\").hide();" +
                "$(\".footnote\").hide();");
            Thread.Sleep(500);
            var rowAmount = await webView21.CoreWebView2.ExecuteScriptAsync(
                "var o = document.getElementsByClassName(\"kurlarTablo\");" +
                "var rows = o[0].lastChild.childNodes;" +
                "rows.length"
                );
            dataGridView1.Rows.Clear();
            export.Enabled= false;
            bool colorControl = true; //dataGridView renk formatlamasında önceki satırın rengine göre renk değişikliği yapmak için bool değer
            if (rowAmount != "null")    //sayfaya tablo yüklenmiş ise data işlenerek dataGrıdView e aktarılır
            {   
                for (int i = 0; i < int.Parse(rowAmount); i += 2)
                {
                    string kurKodu = await webView21.CoreWebView2.ExecuteScriptAsync("var con = rows[" + i + "].getElementsByTagName(\"td\")[0].textContent; con");
                    kurKodu = kurKodu.Replace(" ", "").Replace("\\", "").Replace("r", "").Replace("n", "").Replace("t", ""); // webSitesi üzerinden gelen kur kodu strınginin gereksiz karakterlerden arındırılması
                    string birim = await webView21.CoreWebView2.ExecuteScriptAsync("rows[" + i + "].getElementsByTagName(\"td\")[1].textContent");
                    string dovizCinsi = await webView21.CoreWebView2.ExecuteScriptAsync("rows[" + i + "].getElementsByTagName(\"td\")[2].textContent");
                    string dovizAlis = await webView21.CoreWebView2.ExecuteScriptAsync("rows[" + i + "].getElementsByTagName(\"td\")[3].textContent");
                    string dovizSatis = await webView21.CoreWebView2.ExecuteScriptAsync("rows[" + i + "].getElementsByTagName(\"td\")[4].textContent");
                    string elektifAlis = await webView21.CoreWebView2.ExecuteScriptAsync("rows[" + i + "].getElementsByTagName(\"td\")[5].textContent");
                    string elektifSatis = await webView21.CoreWebView2.ExecuteScriptAsync("rows[" + i + "].getElementsByTagName(\"td\")[6].textContent");
                    dataGridView1.Rows.Add(new[] {  kurKodu.Substring(1,kurKodu.Length - 2),
                                                birim.Substring(1,birim.Length - 2),
                                                dovizCinsi.Substring(1,dovizCinsi.Length - 2),
                                                dovizAlis.Substring(1,dovizAlis.Length - 2),
                                                dovizSatis.Substring(1,dovizSatis.Length - 2),
                                                elektifAlis.Substring(1,elektifAlis.Length - 2),
                                                elektifSatis.Substring(1,elektifSatis.Length - 2)});
                    if (colorControl)
                    {
                        dataGridView1.Rows[i / 2].DefaultCellStyle.BackColor = Color.LightGray;
                        colorControl = false;
                    }
                    else colorControl = true;
                }
                export.Enabled = true; // data yüklendikten sonra excel dışa aktar butonu aktif edilir
            }
            else // sayfaya tablo yüklenmemişse TCBM sisteminin de verdiği hata mesajı messageBox olarak gösterilir
            {
                MessageBox.Show("Resmi tatil, hafta sonu ve yarım iş günü çalışılan günlerde gösterge niteliginde kur bilgisi yayımlanmamaktadır.");
            }
            submit.Enabled = true; // data yükleme işlemi bittiği için yeni data gösterme talebi için buton aktif edilmiştir 
        }

        public bool exportExcel(string fileName)
        {
            bool isExported = true;
            using(IXLWorkbook workbook = new XLWorkbook())
            {
                workbook.AddWorksheet("Sheet 1").FirstCell().InsertData(fileName);
                bool colorControl = true; // excel dosyasının renk formatlaması için önceki satırın renk kontrolünü sağlayan bool değer
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    workbook.Worksheet("Sheet 1").Cell(1, i).SetValue<string>(dataGridView1.Columns[i - 1].HeaderText); //sütun açıklamalarının aktarılması
                }
                for (int i = 0; i < dataGridView1.Rows.Count-1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)//dataGridView içerisindeki data excel dosyasına kopyalanır
                    {
                        workbook.Worksheet("Sheet 1").Cell(i+2, j+1).SetValue<string>(dataGridView1.Rows[i].Cells[j].Value.ToString());
                        if (colorControl)
                            workbook.Worksheet("Sheet 1").Cell(i + 2, j + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    if(colorControl)
                        colorControl= false;
                    else colorControl= true;
                }
                workbook.Worksheet("Sheet 1").Columns().AdjustToContents(); //sütun genişlikleri veriye göre ayarlanır
                workbook.Worksheet("Sheet 1").Range(1, 1, dataGridView1.RowCount, dataGridView1.Columns.Count).Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                workbook.Worksheet("Sheet 1").Range(1, 1, dataGridView1.RowCount, dataGridView1.Columns.Count).Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
                workbook.Worksheet("Sheet 1").Range(1, 1, 1, 7).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                workbook.SaveAs(fileName);
                isExported = true;
            }
            return isExported;
        }

        private void export_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.OverwritePrompt = false;
            save.CreatePrompt = true;
            save.InitialDirectory = @"D:\";
            save.Title = "Excel Dosyaları";
            save.DefaultExt = "xlsx";
            save.Filter = "xlsx Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar(*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                if (exportExcel(save.FileName))
                {
                    MessageBox.Show("Dosya dışa aktarıldı.");
                };
            }
            
        }
    }
}