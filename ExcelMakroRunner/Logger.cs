using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelMakroRunner
{
    public static class Logger
    {
        // Log dosyalarının kaydedileceği klasör yolu.
        private static string? logPath;

        // Log klasörü tanımlanır.
        public static void Initialize(string path)
        {
            logPath = path;

            // Klasör yoksa oluşturulur.
            Directory.CreateDirectory(logPath);
        }

        // Gelen mesajı hem konsola hem de log dosyasına yazar.
        public static void Log(string message)
        {
            // Tarih ve saat bilgisi alınır.
            string time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            // Yazılacak log satırı oluşturulur.
            string logMessage = $"{time} - {message}";

            // Mesaj konsola yazdırılır.
            Console.WriteLine(logMessage);

            // Logger başlatılmamışsa hata oluşturulur.
            if (string.IsNullOrWhiteSpace(logPath))
            {
                throw new InvalidOperationException("Logger başlatılmadı.");
            }

            // Günlük log dosyasının yolu oluşturulur.
            string logFile = Path.Combine(
                logPath,
                $"{DateTime.Now:yyyy-MM-dd}.log"
            );

            // Log mesajı dosyanın sonuna eklenir.
            File.AppendAllText(
                logFile,
                logMessage + Environment.NewLine
            );
        }
    }
}