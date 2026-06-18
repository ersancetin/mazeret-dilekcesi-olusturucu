# Mazeret Dilekçesi Oluşturucu

UYAP duruşma listesinden aynı güne denk gelen duruşmalar için otomatik mazeret dilekçesi taslağı oluşturan, tamamen tarayıcıda çalışan bir araç.

## Özellikler

- **UYAP Excel desteği** — Avukat Portal'dan indirilen duruşma listesini doğrudan yükleyin
- **Otomatik tespit** — Aynı güne denk gelen duruşmaları otomatik bulur
- **UDF formatı** — UYAP editöründe açılabilen `.udf` dosyaları oluşturur
- **Toplu indirme** — Gün bazlı veya tümünü tek ZIP olarak indirin
- **Arama** — Birim veya dosya numarasına göre filtreleme
- **Gizlilik** — Hiçbir veri sunucuya gönderilmez, tüm işlemler tarayıcıda gerçekleşir
- **Mobil uyumlu** — Telefon ve tabletten de kullanılabilir

## Kullanım

1. UYAP Avukat Portal'dan duruşma listenizi Excel olarak indirin
2. Dosyayı sürükleyip bırakın veya tıklayarak yükleyin
3. Avukat bilgilerinizi girin (ad soyad, baro, sicil no, adres)
4. Oluşturulan dilekçeleri gün bazlı veya tek tek indirin
5. İndirilen `.udf` dosyalarını UYAP editöründe açıp kontrol ettikten sonra sisteme yükleyin

## Beklenen Excel Formatı

| Sütun | İçerik |
|-------|--------|
| A | Birim (Mahkeme adı) |
| B | Dosya No |
| D | Duruşma Tarihi |

## Canlı Demo

[https://ersancetin.github.io/mazeret-dilekcesi-olusturucu](https://ersancetin.github.io/mazeret-dilekcesi-olusturucu)

## Sorumluluk Reddi

Bu araç mazeret dilekçesi **taslağı** oluşturur. Oluşturulan dilekçelerin içeriği kullanıcı tarafından kontrol edilmelidir. Kullanımdan doğabilecek her türlü sorumluluk kullanıcıya aittir.

## Kullanılan Kütüphaneler

- [SheetJS](https://sheetjs.com/) — Excel dosya okuma
- [JSZip](https://stuk.github.io/jszip/) — ZIP dosya oluşturma
- [FileSaver.js](https://github.com/nicolo-ribaudo/FileSaver.js) — Dosya indirme

UDF dosya formatının anlaşılmasında [UDF-Toolkit](https://github.com/AhmetSBulbul/UDF-Toolkit) projesinden referans alınmıştır.

## Lisans

[MIT](LICENSE)

## İletişim

**Ersan Cetin** — [GitHub](https://github.com/ersancetin) · [LinkedIn](https://linkedin.com/in/ersan-cetin)
