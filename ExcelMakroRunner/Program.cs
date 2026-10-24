using ExcelMakroRunner;
using Microsoft.Extensions.Configuration;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel; // Excel kütüphanesine kısa isim verildi.

// finally bloğunda erişilebilmeleri için değişkenler try bloklarının dışında tanımlandı.
string excelPath;
string logPath;
string[]? macros = null;
Excel.Application? excelApp = null;
Excel.Workbooks? workbooks = null;
Excel.Workbook? workbook = null;

// settings.ini dosyasını okumak için yapılandırma oluşturuldu.
IConfiguration configuration = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory) // Programın çalıştığı klasör seçildi.
    .AddIniFile("settings.ini", optional: false, reloadOnChange: true) // settings.ini dosyası eklendi.
    .Build(); // Yapılandırma hazırlandı.

// 1. Aşama: Ayarlar Excel açılmadan önce okunur ve doğrulanır.
try
{
    // Hataların dosyaya yazılabilmesi için önce log klasörü okunur ve Logger başlatılır.
    logPath = configuration["General:LogPath"]
        ?? throw new Exception("settings.ini dosyasında LogPath anahtarı bulunamadı.");
    Logger.Initialize(logPath);

    // Çalıştırılacak makroların listesi okunur.
    macros = configuration
        .GetSection("Macros") // Macros bölümü seçildi.
        .Get<string[]>() // Liste string dizisine dönüştürüldü.
        ?? throw new Exception("settings.ini dosyasında Macros anahtarı bulunamadı.");

    // Açılacak Excel dosyasının yolu okunur.
    excelPath = configuration["General:ExcelPath"]
        ?? throw new Exception("settings.ini dosyasında ExcelPath anahtarı bulunamadı.");
}
catch (Exception ex)
{
    // Logger başlatılamamış olabileceği için hata konsola yazılır.
    Console.WriteLine("Ayarlar okunurken hata oluştu.");
    Console.WriteLine(ex.Message);

    // Görev Zamanlayıcı'nın çalıştırmayı başarısız görmesi için çıkış kodu 1 yapılır.
    Environment.ExitCode = 1;

    // Excel hiç açılmadan program sonlandırılır.
    return;
}

// 2. Aşama: Excel açılır ve makrolar çalıştırılır.
try
{
    // Arka planda yeni bir Excel uygulaması başlatıldı.
    excelApp = new Excel.Application();

    // Excel kullanıcıya görünmeden çalışır.
    excelApp.Visible = false;
    excelApp.DisplayAlerts = false;
    // "İki nokta" kuralı: Workbooks koleksiyonu serbest bırakılabilmesi için ayrı değişkende tutulur.
    workbooks = excelApp.Workbooks;

    // Excel dosyası açıldı.
    workbook = workbooks.Open(excelPath);
    Logger.Log("Excel açıldı.");

    // Tüm makrolar sırayla çalıştırılır.
    foreach (string macro in macros)
    {
        Logger.Log($"Makro çalıştırılıyor: {macro}");
        excelApp.Run(macro);
        Logger.Log($"Makro çalıştırıldı: {macro}");
    }
}
catch (Exception ex)
{
    // Görev Zamanlayıcı'nın çalıştırmayı başarısız görmesi için çıkış kodu 1 yapılır.
    Environment.ExitCode = 1;

    // Hata, satır bilgisiyle birlikte loglanır.
    Logger.Log("Hata oluştu.");
    Logger.Log(ex.ToString());
}
finally
{
    // Önce Excel'e kapanması söylenir.
    // Kaynak dosyanın değişmemesi için çalışma kitabı kaydedilmeden kapatılır (bilinçli karar).
    if (workbook != null)
    {
        workbook.Close(false);
    }
    if (excelApp != null)
    {
        excelApp.Quit();
    }

    // Ardından COM nesneleri, oluşturulma sırasının tersine serbest bırakılır.
    // Böylece arka planda EXCEL.EXE süreci açık kalmaz.
    if (workbook != null)
    {
        Marshal.ReleaseComObject(workbook);
    }
    if (workbooks != null)
    {
        Marshal.ReleaseComObject(workbooks);
    }
    if (excelApp != null)
    {
        Marshal.ReleaseComObject(excelApp);
    }

    Logger.Log("Excel kapatıldı.");
}