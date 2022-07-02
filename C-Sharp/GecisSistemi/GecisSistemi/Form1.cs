using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;       
namespace GecisSistemi
{
    public partial class Form1 : Form
    {
        SerialPort serialPort;
        bool serialPortDurum = false;
        List<Kisi> kisiListesi = new List<Kisi>();
        List<Kisi> girisYapanListesi = new List<Kisi>();
        Kisi anlikKisi;
        string ExcelDosyaYolu;
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelDosyaKontrol();
            buttonOK.Enabled = false;
            ButonPasif();
            KisileriIceriAktar();
            serialPort = new SerialPort();
            string[] ports = SerialPort.GetPortNames();
            foreach(string port in ports){
                comboBoxPort.Items.Add(port);
            }
        }
        private void ButonAktif()
        {
            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = true;
            button13.Enabled = true;
            button14.Enabled = true;
            button15.Enabled = true;
            button16.Enabled = true;
            button17.Enabled = true;
            button18.Enabled = true;
            button19.Enabled = true;
            button20.Enabled = true;
            button21.Enabled = true;
            button22.Enabled = true;
            button23.Enabled = true;
            button24.Enabled = true;
            button25.Enabled = true;
            button26.Enabled = true;
            button27.Enabled = true;
        }
        private void ButonPasif()
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            button4.Enabled = false;
            button5.Enabled = false;
            button6.Enabled = false;
            button7.Enabled = false;
            button8.Enabled = false;
            button9.Enabled = false;
            button10.Enabled = false;
            button11.Enabled = false;
            button12.Enabled = false;
            button13.Enabled = false;
            button14.Enabled = false;
            button15.Enabled = false;
            button16.Enabled = false;
            button17.Enabled = false;
            button18.Enabled = false;
            button19.Enabled = false;
            button20.Enabled = false;
            button21.Enabled = false;
            button22.Enabled = false;
            button23.Enabled = false;
            button24.Enabled = false;
            button25.Enabled = false;
            button26.Enabled = false;
            button27.Enabled = false;
        }
        private void KisileriIceriAktar()
        {
            string url = Directory.GetCurrentDirectory() + @"\Kisiler.csv";
            if (!File.Exists(url))
            {
                MessageBox.Show("Kisiler.csv dosyası bulunamadı. Lütfen oluşturup tekrar deneyiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.ExitThread();
            }
            StreamReader sr = new StreamReader(url, Encoding.UTF8);
            string line = sr.ReadLine();
            string ad,soyad,kartId;
            while (true){
                line = sr.ReadLine();
                if (line == null){
                    break;
                }
                var list = line.Split(';');
                try{
                    ad = list[0].Replace("\"", "");
                    soyad = list[1].Replace("\"", "");
                    kartId = list[2].Replace("\"", "");
                    kisiListesi.Add(new Kisi(ad,soyad, kartId));
                }
                catch (Exception){
                    MessageBox.Show("CSV Dosyası Hatali! Lütfen düzenleyip tekrar deneyiniz.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.ExitThread();
                }
            }
            sr.Close();
        }
        private void ExcelDosyaKontrol()
        {
            string konum = Directory.GetCurrentDirectory() + @"\Veriler\";
            if (!Directory.Exists(konum))
            {
                Directory.CreateDirectory(konum);
            }
            string tarih = DateTime.Now.ToString("dd.MM.yyyy");
            string dosya = tarih + ".xlsx";
            ExcelDosyaYolu = konum + dosya;
            int i = 1;
            while (File.Exists(ExcelDosyaYolu))
            {
                ExcelDosyaYolu = konum + tarih + "(" + i + ")" + ".xlsx";
                i++;
            }
        }
        private void buttonPortBaglan_Click(object sender, EventArgs e){
            if (serialPortDurum == false){
                if (comboBoxPort.SelectedItem == null){
                    MessageBox.Show("Lütfen Port Seçiniz!");
                }
                else{
                    serialPort.PortName = comboBoxPort.SelectedItem.ToString();
                    serialPort.BaudRate = 9600;
                    serialPort.Open();
                    serialPortDurum = true;
                    buttonPortBaglan.Text = "Bağlantı Kapat";
                    MessageBox.Show("Bağlantı açıldı!", "Bilgi",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    serialPort.DataReceived += VeriGeldi;
                }
            }
            else if (serialPortDurum == true){
                if (serialPort != null && serialPort.IsOpen)
                {
                    while (!(serialPort.BytesToRead == 0 && serialPort.BytesToWrite == 0))
                    {
                        serialPort.DiscardInBuffer();
                        serialPort.DiscardOutBuffer();
                    }
                }
                ButonPasif();
                serialPort.Close();
                serialPortDurum = false;
                buttonPortBaglan.Text = "Bağlantıyı Aç";
                MessageBox.Show("Bağlantı kapandı!","Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void VeriGeldi(object sender, SerialDataReceivedEventArgs e)
        {
            var serialCihaz = sender as SerialPort;
            string veri = serialCihaz.ReadLine();
            veri = veri.Replace("\n", "").Replace("\r", "");
            if (veri[0] == '_' && veri[veri.Length - 1] == '_')
            {
                string kartId = "";
                for (int i = 1; i < veri.Length - 1; i++)
                {
                    kartId += veri[i];
                }
                foreach (Kisi kisi in kisiListesi)
                {
                    if(kisi.KartId == kartId && kisi.GirdiMi==false){
                        anlikKisi = kisi;
                        Invoke(new EventHandler(Giris));
                    }
                    else if(kisi.KartId == kartId && kisi.GirdiMi ==true){
                        anlikKisi = kisi;
                        Invoke(new EventHandler(Cikis));
                    }
                }
            }
        }
        private void Giris(object sender, EventArgs e)
        {
            ButonAktif();
            DateTime dt = DateTime.Now;
            labelSonİslem.Text = "Son İşlem : " + anlikKisi.Ad + " Giriş Yaptı.";
            labelSonIslemTarih.Text = "Son İşlem Tarihi : " + dt.ToString("dd.MM.yyyy HH:mm:ss");
            anlikKisi.GirisTarihi = dt.ToString("dd.MM.yyyy HH:mm:ss");
            anlikKisi.GirisDateTime = dt;
            anlikKisi.GirdiMi = true;
            Kisi tempKisi = new Kisi(anlikKisi.Ad, anlikKisi.Soyad, anlikKisi.KartId);
            tempKisi.GirisTarihi = anlikKisi.GirisTarihi;
            tempKisi.GirdiMi = anlikKisi.GirdiMi;
            tempKisi.GirisDateTime = anlikKisi.GirisDateTime;
            tempKisi.Durum = anlikKisi.Durum;
            girisYapanListesi.Add(tempKisi);
            ExcelYaz();
        }
        private void Cikis(object sender, EventArgs e)
        {
            ButonPasif();
            DateTime dt = DateTime.Now;
            labelSonİslem.Text = "Son İşlem : "+anlikKisi.Ad+" Çıkış Yaptı.";
            labelSonIslemTarih.Text = "Son İşlem Tarihi : "+dt.ToString("dd.MM.yyyy HH:mm:ss");
            anlikKisi.CikisTarihi = dt.ToString("dd.MM.yyyy HH:mm:ss");
            anlikKisi.CikisDateTime = dt;
            anlikKisi.GirdiMi = false;
            TimeSpan ts = anlikKisi.CikisDateTime - anlikKisi.GirisDateTime;
            double dakika = ts.TotalMinutes;
            anlikKisi.SureFarkDakika = string.Format("{0:N2}", dakika);
            Kisi tempKisi = new Kisi(anlikKisi.Ad, anlikKisi.Soyad, anlikKisi.KartId);
            tempKisi.GirisTarihi = anlikKisi.GirisTarihi;
            tempKisi.CikisTarihi = anlikKisi.CikisTarihi;
            tempKisi.SureFarkDakika = anlikKisi.SureFarkDakika;
            tempKisi.GirdiMi = anlikKisi.GirdiMi;
            tempKisi.CikisDateTime = anlikKisi.CikisDateTime;
            tempKisi.GirisDateTime = anlikKisi.GirisDateTime;
            tempKisi.Durum = anlikKisi.Durum;
            for (int i = 0; i < girisYapanListesi.Count; i++)
            {
                if (girisYapanListesi[i].GirisTarihi == tempKisi.GirisTarihi)
                {
                    girisYapanListesi[i] = tempKisi;
                }
            }
            ExcelYaz();
            anlikKisi = null;
        }
        private void ExcelYaz()
        {
            Excel.Application ExcelUygulama;
            Excel.Workbook CalismaKitabi;
            Excel.Worksheet CalismaSayfasi;
            ExcelUygulama = new Excel.Application();
            object misValue = System.Reflection.Missing.Value;
            CalismaKitabi = ExcelUygulama.Workbooks.Add(misValue);
            CalismaSayfasi = (Excel.Worksheet)CalismaKitabi.Worksheets.get_Item(1);
            CalismaSayfasi.Name = "Veriler";
            CalismaSayfasi.Cells[1, 1] = "Ad";
            CalismaSayfasi.Cells[1, 2] = "Soyad";
            CalismaSayfasi.Cells[1, 3] = "KartId";
            CalismaSayfasi.Cells[1, 4] = "Giris Tarihi";
            CalismaSayfasi.Cells[1, 5] = "Çıkış Tarihi";
            CalismaSayfasi.Cells[1, 6] = "Kaldığı Süre Dakika";
            CalismaSayfasi.Cells[1, 7] = "Durum";
            CalismaSayfasi.Rows.Cells.Style.Font.Size = 20;
            CalismaSayfasi.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            var Aralik = CalismaSayfasi.Range[CalismaSayfasi.Cells[1, 1], CalismaSayfasi.Cells[1, 7]];
            Aralik.Interior.Color = Excel.XlRgbColor.rgbOrange;
            int i = 2;
            foreach (Kisi kisi in girisYapanListesi)
            {
                CalismaSayfasi.Cells[i, 1] = kisi.Ad;
                CalismaSayfasi.Cells[i, 2] = kisi.Soyad;
                CalismaSayfasi.Cells[i, 3] = "'" + kisi.KartId;
                CalismaSayfasi.Cells[i, 4] = kisi.GirisTarihi;
                CalismaSayfasi.Cells[i, 5] = kisi.CikisTarihi;
                CalismaSayfasi.Cells[i, 6] = kisi.SureFarkDakika;
                CalismaSayfasi.Cells[i, 7] = kisi.Durum;
                i++;
            }
            CalismaSayfasi.Columns.AutoFit();
            ExcelUygulama.Visible = false;
            ExcelUygulama.DisplayAlerts = false;
            try
            {
                CalismaKitabi.SaveAs(ExcelDosyaYolu);
                CalismaKitabi.Close(true);
            }
            catch (Exception)
            {
                MessageBox.Show("Excel dosyası kaydedilemedi. Dosya açık olabilir. Lütfen dosyayı kapatınız.", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            ExcelUygulama.Quit();
        }
        private void buttonPortYenile_Click(object sender, EventArgs e){
            comboBoxPort.Items.Clear();
            comboBoxPort.Text = "";
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                comboBoxPort.Items.Add(port);
            }
        }
        private void button_Click(object sender, EventArgs e)
        {
            labelSecilenDurum.Text = (sender as Button).Text;
            buttonOK.Enabled = true;
        }
        private void buttonOK_Click(object sender, EventArgs e)
        {
            anlikKisi.Durum = labelSecilenDurum.Text;
            for (int i = 0; i < girisYapanListesi.Count; i++)
            {
                if (girisYapanListesi[i].GirisTarihi == anlikKisi.GirisTarihi)
                {
                    girisYapanListesi[i].Durum = labelSecilenDurum.Text;
                    girisYapanListesi[i].CikisTarihi = ""; ;
                    girisYapanListesi[i].SureFarkDakika = ""; 
                }
            }
            labelSecilenDurum.Text = "Durum Seçiniz!";
            ButonPasif();
            buttonOK.Enabled = false;
            ExcelYaz();
        }
        private void buttonKaydet_Click(object sender, EventArgs e)
        {
            ExcelYaz();
        }
    }
}




