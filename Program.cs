using System;
using System.Collections.Generic; //listeler için
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//sonradan eklenenler
using System.Data.SqlClient;  //sql işlemleri için
using System.Net; //web hatalarını yakalar
using ClosedXML.Excel;
using System.Xml;
using System.IO; //masaüstü yolu bulmak için
using System.Configuration; //config okumak için

namespace TcmbKurCekici
{
    internal class Program
    {
        //config den bilgilere erişme
        // "BenimDbBaglantim" yazan yer, App.config dosyasında verdiğimiz name değeriyle birebir aynı olmalı.
        static string connectionString = ConfigurationManager.ConnectionStrings["BenimDbBaglantim"].ConnectionString;

        static void Main(string[] args)
        {
            /* MB APİ SİNE ULAŞMA VE KURLARI ÇEKME
             * Console.WriteLine("TC Merkez Bankası Api ile kur getirme programı başlıyor...");

             //gidilecek adresi belirt
             string url = "https://www.tcmb.gov.tr/kurlar/202608/10082026.xml";
             Console.WriteLine($"Gidilen adres: {url}");

             //xml okuyucu nesneyi oluşturup, adrese yolluyoruz
             XmlDocument xmlDoc = new XmlDocument();
             xmlDoc.Load(url);

             //dolar değerlerini xml içindeki XPath ile buluyor
             string dolarStr = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='USD']/BanknoteSelling").InnerText;

             //euro değerini çekiyoruz
             string euroStr = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='EUR']/BanknoteSelling").InnerText;

             //ekrana yaz
             Console.WriteLine($"\n Başarılı. Kurlar çekildi. ");
             Console.WriteLine($"1 Dolar: {dolarStr} TL");
             Console.WriteLine($"1 Euro: {euroStr} TL");
            */

            //Hafta sonlarına (örneğin 15-16 Ağustos) denk geldiğinde program çökmeyecek, "tatil, 1 gün geriye gidiliyor" diyecek ve Cuma gününün kurunu alıp yola devam edecek.
            //Bugün 25 Ağustos olduğu için ve saat henüz 15:30'u geçmediği için de 25'i sorduğumuzda 24'ünün kurunu getirecek.

            Console.WriteLine("10 - 25 Ağustos Arası Kur Çekme İşlemi Başlıyor...\n");

            // Başlangıç ve bitiş tarihlerimizi belirliyoruz
            DateTime baslangicTarihi = new DateTime(2026, 8, 10);
            DateTime bitisTarihi = new DateTime(2026, 8, 25);

            // Tüm kurları hafızada tutup en son Excel'e basmak için bir liste oluşturuyoruz
            List<KurVerisi> tumKurlar = new List<KurVerisi>();

            // for döngüsü ile başlangıçtan bitişe kadar gün gün ilerliyoruz (tarih.AddDays(1) = 1 gün ekle)
            for (DateTime tarih = baslangicTarihi; tarih <= bitisTarihi; tarih = tarih.AddDays(1))
            {
                Console.WriteLine("--------------------------------------------------");
                Console.WriteLine($"\n İşlenen Tarih: {tarih.ToString("dd.MM.yyyy")}");

                // Birazdan aşağıda yazacağımız metoda tarihi gönderip kuru alıyoruz
                KurVerisi cekilenKur = KurGetir(tarih);

                if (cekilenKur != null)
                {
                    // 2. Veritabanına Yaz
                    VeritabaninaKaydet(cekilenKur);

                    // 3. Excel için listeye ekle
                    tumKurlar.Add(cekilenKur);
                }

                // Console.WriteLine($"Bulunan Dolar: {cekilenKur.Dolar} | Bulunan Euro: {cekilenKur.Euro}");
                // Console.WriteLine($"Baz Alınan TCMB Tarihi: {cekilenKur.GecerliTarih.ToString("dd.MM.yyyy")}");
            }

            Console.WriteLine("\n--------------------------------------------------");
            Console.WriteLine("Tüm tarihler tarandı. Excel dosyası oluşturuluyor...");

            // 4. Bütün işlemler bitince Excel dosyasını oluştur
            ExcelOlustur(tumKurlar);

            Console.WriteLine("\nİşlem Başarıyla Tamamlandı! Kapatmak için bir tuşa basın.");
            Console.ReadLine();
        }

        static KurVerisi KurGetir(DateTime istenenTarih)
        {
            DateTime kontrolEdilenTarih = istenenTarih;

            //bugünse ve 15.30 dan önceyse kuralı
            if (istenenTarih.Date == DateTime.Today.Date && DateTime.Now.TimeOfDay < new TimeSpan(15, 30, 0))
            {
                kontrolEdilenTarih = kontrolEdilenTarih.AddDays(-1); //1 gün geriye gider
                Console.WriteLine("-> Uyarı: Saat 15.30'u geçmediği için düne ait kur verileri gelecektir.");
            }

            //hafta sonu veya tatilse kuralı
            while (true)
            {
                try
                {
                    //tarihi linkin istediği formata çeviriyoruz
                    string yilAy = kontrolEdilenTarih.ToString("yyyyMM");
                    string gunAyYil = kontrolEdilenTarih.ToString("ddMMyyyy");

                    string url = $"https://www.tcmb.gov.tr/kurlar/{yilAy}/{gunAyYil}.xml";

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(url); //tatil günü ise xml yayınlanmamıştır ve hata fırlatır, catch e gider

                    //hata yoksa devam eder 
                    string dolarStr = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='USD']/BanknoteSelling").InnerText;
                    string euroStr = xmlDoc.SelectSingleNode("Tarih_Date/Currency[@Kod='EUR']/BanknoteSelling").InnerText;

                    //noktaları virgüle çevirme (decimal)
                    decimal dolar = Convert.ToDecimal(dolarStr.Replace('.', ','));
                    decimal euro = Convert.ToDecimal(euroStr.Replace('.', ','));

                    //verileri paket yapıp geri gönderiyoruz
                    return new KurVerisi
                    {
                        SorgulananTarih = istenenTarih,
                        GecerliTarih = kontrolEdilenTarih,
                        Dolar = dolar,
                        Euro = euro

                    };
                }

                catch (WebException)
                {
                    Console.WriteLine($"-> {kontrolEdilenTarih.ToString("dd.MM.yyyy")} tatil, 1 gün geriye gidiyor...");
                    kontrolEdilenTarih = kontrolEdilenTarih.AddDays(-1); //1 gün geri git
                }
            }

        }

        static void VeritabaninaKaydet(KurVerisi kur)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();

                    // SQL Sorgumuzu güncelledik: "Eğer bu tarih tablomuzda yoksa (IF NOT EXISTS), o zaman INSERT yap."
                    string query = @"
                    IF NOT EXISTS (SELECT 1 FROM KurTablosu WHERE Tarih = @Tarih)
                     BEGIN
                    INSERT INTO KurTablosu (Tarih, Dolar, Euro) VALUES (@Tarih, @Dolar, @Euro)
                    END";

                    using (SqlCommand cmd = new SqlCommand (query, con))
                    {
                        cmd.Parameters.AddWithValue("@Tarih", kur.SorgulananTarih);
                        cmd.Parameters.AddWithValue("@Dolar", kur.Dolar);
                        cmd.Parameters.AddWithValue("@Euro", kur.Euro);

                        // ExecuteNonQuery bize işlemden etkilenen satır sayısını (int olarak) döndürür.
                        int etkilenenSatir = cmd.ExecuteNonQuery();

                        // Eğer etkilenen satır 0'dan büyükse kayıt eklenmiştir, değilse atlanmıştır.
                        if (etkilenenSatir > 0)
                        {
                            Console.WriteLine("-> Veritabanına yeni kayıt olarak eklendi.");
                        }
                        else
                        {
                            Console.WriteLine("-> Bu tarih zaten veritabanında var, ekleme atlandı.");
                        }

                    }
                }
               
            }
            catch (Exception ex)
            {
                Console.WriteLine($"-> SQL Hatası: {ex.Message}");
            }
        }

        static void ExcelOlustur(List<KurVerisi> kurlar)
        {
            //masaüstü yolunu bulup dosya oluştur
            string masaUstuYolu = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string dosyaYolu = $@"{masaUstuYolu}\Tcmb_Kurlar_10_25_Agustos.xlsx";

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Kurlar");

                //1. Satır başlıklardan oluşuyor
                worksheet.Cell(1, 1).Value = "İstenen Tarih";
                worksheet.Cell(1, 2).Value = "TCMB Kayıt Tarihi";
                worksheet.Cell(1, 3).Value = "Dolar (USD)";
                worksheet.Cell(1, 4).Value = "Euro (EUR)";

                // Başlıkları kalın (bold) yapalım 
                worksheet.Range("A1:D1").Style.Font.Bold = true;

                //2. Satırdan itibaren verileri yazıyor
                int satir = 2;
                foreach(var kur in kurlar)
                {
                    worksheet.Cell(satir, 1).Value = kur.SorgulananTarih;
                    worksheet.Cell(satir, 2).Value = kur.GecerliTarih;
                    worksheet.Cell(satir, 3).Value = kur.Dolar;
                    worksheet.Cell(satir, 4).Value = kur.Euro;
                    satir++;
                }

                // Sütun genişliklerini içindeki yazıya göre otomatik ayarla
                worksheet.Columns().AdjustToContents();

                //Dosyayı kaydet
                workbook.SaveAs(dosyaYolu);

            }

            Console.WriteLine($"-> Excel dosyası masaüstüne kaydedildi: Tcmb_Kurlar_10_25_Agustos.xlsx");
        }

        class KurVerisi
        {
            public DateTime SorgulananTarih { get; set; }
            public DateTime GecerliTarih { get; set; }
            public decimal Dolar { get; set; }
            public decimal Euro { get; set; }
        }


    }
}
