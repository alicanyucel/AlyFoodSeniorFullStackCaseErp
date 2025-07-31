using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using System;
using System.Diagnostics;
using System.IO;

class Program
{
    static void Main()
    {
        // Yeni belge oluştur
        var document = new Document();
        var section = document.AddSection();

        // Başlık
        var title = section.AddParagraph("GENEL ÜRETİM ÇALIŞMA PLANI");
        title.Format.Font.Size = 16;
        title.Format.Font.Bold = true;
        title.Format.Font.Name = "Verdana"; // Türkçe karakter uyumlu
        title.Format.Alignment = ParagraphAlignment.Center;
        section.AddParagraph();

        // Bölümleri ekle
        AddSection(document, "1. Vardiya Planlaması",
            "• Vardiya 1: 06:00 – 14:00\n• Vardiya 2: 14:00 – 22:00\n• Vardiya 3: 22:00 – 06:00 (isteğe bağlı)\n" +
            "• Her vardiyada minimum 1 süpervizör ve yeterli sayıda operatör yer almalı.\n• Vardiya devir teslimlerinde üretim bilgisi eksiksiz aktarılmalı.");

        AddSection(document, "2. Üretim Süreci Aşamaları",
            "A. Hazırlık:\n- Giriş kontrolleri (malzeme uygunluğu, makina çalışırlığı)\n- Ürün reçetesi ve talimat kontrolü\n" +
            "- Üretim hattı temizliği ve hijyen kontrolü\n\n" +
            "B. Üretim Takibi:\n- Üretim ekranındaki veriler saatlik olarak izlenmeli:\n" +
            "  Planlanan – Üretilen – Kalan – % Tamamlanma\n- Hedef: %90 üzeri üretim tamamlama oranı\n\n" +
            "C. Kalite Kontrol:\n- Her saat başı kalite kontrol yapılmalı\n- Hatalı ürünler izole edilmeli, sebep analizi yapılmalı\n- Gerekirse teknik müdahale yapılmalı");

        AddSection(document, "3. Temizlik ve 5S Uygulamaları",
            "- Günlük temizlik vardiya bitiminde\n- Haftalık derin temizlik (makineler, zemin, raflar)\n" +
            "- 5S uygulaması:\n  • Seiri (Sınıflandır)\n  • Seiton (Düzenle)\n  • Seiso (Temizle)\n  • Seiketsu (Standartlaştır)\n  • Shitsuke (Sürdür)");

        AddSection(document, "4. Depolama ve Lojistik",
            "- Paletli ürünler etiketli şekilde istiflenmeli\n- Her ürünün üretim tarihi ve lot numarası kontrol edilmeli\n- Sevkiyat planı önceden hazırlanmalı");

        AddSection(document, "5. İş Güvenliği",
            "- Her çalışan:\n  • Baret, iş ayakkabısı, bone ve eldiven kullanmalı\n  • Sarı güvenlik alanlarını ihlal etmemeli\n" +
            "- Acil çıkış yolları ve yangın söndürücüler açık olmalı");

        AddSection(document, "6. Performans ve Raporlama",
            "- Günlük:\n  • Üretim verisi → Süpervizör raporu\n  • Arıza süresi, duruş nedenleri → Teknik rapor\n" +
            "- Haftalık:\n  • KPI analizi: Hedef karşılama, fire oranı, verimlilik");

        AddSection(document, "7. Hedefler ve Motivasyon",
            "- Günlük hedef üretim: Örneğin 300 kasa\n- Haftalık ödüllendirme sistemi (ekip bazlı)\n- Eğitim günleri: Ayda 1 kez süreç verimliliği eğitimi");

        // PDF olarak kaydet
        var renderer = new PdfDocumentRenderer(true)
        {
            Document = document
        };
        renderer.RenderDocument();

        // Masaüstüne kaydet
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        string filePath = Path.Combine(desktopPath, "Genel_Uretim_Calisma_Plani.pdf");

        renderer.PdfDocument.Save(filePath);

        Console.WriteLine("PDF başarıyla oluşturuldu!\nDosya yolu: " + filePath);

        // Otomatik aç (isteğe bağlı)
        Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
    }

    static void AddSection(Document doc, string title, string content)
    {
        var section = doc.Sections[doc.Sections.Count - 1];

        var header = section.AddParagraph(title);
        header.Format.Font.Bold = true;
        header.Format.Font.Size = 12;
        header.Format.Font.Name = "Verdana";
        header.Format.SpaceBefore = "1cm";

        var body = section.AddParagraph(content);
        body.Format.Font.Size = 10;
        body.Format.Font.Name = "Verdana";
        body.Format.SpaceAfter = "0.5cm";
    }
}
