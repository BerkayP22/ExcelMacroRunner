using Excel = Microsoft.Office.Interop.Excel; // Excel kütüphanesi tanıtıldı, kısa isim olarak Excel verildi.

Excel.Application excelApp = new Excel.Application(); // Windows üzerinde yeni bir Excel uygulaması örneği oluşturuldu.

excelApp.Visible = true; // Excel'in ekranda görünmesi sağlandı. Test için true, otomasyon için false yapılabilir.

string ExcelPath = @"C:\Users\pekta\Desktop\ExcelMakroFile\ExcelMakroFile"; // Açılacak Excel dosyasının yolu belirlendi.

Excel.Workbook workbook = excelApp.Workbooks.Open(ExcelPath); // Belirtilen Excel dosyası açıldı.

Console.WriteLine("Excel açıldı.");

try
{
    // Excel içerisindeki RenkDegistiren1 isimli makroyu çalıştırır.
    excelApp.Run("RenkDegistiren1");
    Console.WriteLine("RenkDegistiren1 çalıştırıldı.");

    // İkinci makroyu çalıştırır.
    excelApp.Run("RenkDegistiren2");
    Console.WriteLine("RenkDegistiren2 çalıştırıldı.");

    // Üçüncü makroyu çalıştırır.
    excelApp.Run("RenkDegistiren3");
    Console.WriteLine("RenkDegistiren3 çalıştırıldı.");

    // Dördüncü ve son makroyu çalıştırır.
    excelApp.Run("RenkDegistiren4");
    Console.WriteLine("RenkDegistiren4 çalıştırıldı.");
}
catch (Exception ex)
{
    // try bloğu içerisinde herhangi bir hata oluşursa program buraya gelir.
    Console.WriteLine("Hata oluştu.");

    // ex nesnesi oluşan hata hakkındaki bilgileri tutar.
    Console.WriteLine(ex.Message);
}
finally
{
    // finally bloğu hata oluşsa da oluşmasa da çalışır.

    // Açılan Excel çalışma kitabını kaydederek kapatması için true verilmeli, test aşamasında false verildi.
    workbook.Close(false);

    // C# tarafından başlatılan Excel uygulamasını kapatır.
    excelApp.Quit();

    Console.WriteLine("Excel kapatıldı.");
}