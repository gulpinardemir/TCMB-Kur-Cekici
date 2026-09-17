# TCMB Kur Çekici (TcmbKurCekici)

Türkiye Cumhuriyet Merkez Bankası (TCMB) açık API'sini kullanarak güncel döviz kurlarını çeken, bu verileri hem MSSQL veritabanına hem de Excel dosyasına raporlayarak kaydeden bir C# konsol uygulamasıdır.

## 🚀 Özellikler

- **TCMB API Entegrasyonu:** Güncel döviz kurlarını anlık olarak XML formatında çeker ve işler.
- **MSSQL Veritabanı Kaydı:** Çekilen verileri ADO.NET aracılığıyla SQL Server'a kaydeder.
- **Mükerrer Kayıt Kontrolü:** Aynı gün ve aynı kur bilgisi için veritabanında mükerrer kayıt oluşmasını engelleyen bir kontrol mekanizması içerir.
- **Excel Raporlama:** `ClosedXML` kütüphanesi kullanılarak çekilen kur bilgileri düzenli bir Excel tablosu olarak dışa aktarılır.

## 🛠️ Kullanılan Teknolojiler

- C# (.NET Framework / Console Application)
- MSSQL Server
- ClosedXML (Excel işlemleri için)
- TCMB EVDS/XML API

## ⚙️ Kurulum ve Kullanım

1. Projeyi bilgisayarınıza indirin (Clone veya ZIP olarak).
2. Projeyi Visual Studio ile açın.
3. `App.config` dosyasını açarak kendi MSSQL veritabanı bağlantı cümlenizi (Connection String) güncelleyin:
   ```xml
   <connectionStrings>
       <add name="DbConnection" connectionString="Server=SUNUCU_ADINIZ;Database=VERITABANI_ADINIZ;User Id=KULLANICI_ADI;Password=SIFRE;" providerName="System.Data.SqlClient" />
   </connectionStrings>
