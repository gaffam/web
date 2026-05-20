# VELIN Desk (Kişisel Yazı Programı)

Bu proje, **sadece kişisel kullanım** için tasarlanmış, çok dilli yazım denetimi ve not yönetimi odaklı masaüstü bir uygulama prototipidir.

## Özellikler
- Türkçe, İngilizce, Fransızca ve İtalyanca yazım denetimi (`language_tool_python`).
- `.md`, `.pdf`, `.docx` dosyalarını açma ve içeriği düzenleyicide gösterme.
- "Pineider Capri" ve "Vélin / Toile ancienne" hissiyatına yakın kağıt dokulu tema.
- Yazarken cursor parlaması ve harf oluşum animasyonu benzeri yumuşak görsel efekt.
- Ajanda + takvim paneli (Moleskine/Hemingway ruhu).
- Zettelkasten not sistemi (benzersiz ID ve çift yönlü bağlantı).
- RSS çekme ve yerel indeksleme.
- Pocket-benzeri “Sonra Oku” listesi.
- Kendi cloud’una (WebDAV) bağlanıp notları senkronlama.
- AI destekli yeniden yazım/özet önerileri (OpenAI API anahtarıyla).
- Windows için tıkla-kurulum mantığına uygun `exe` paketlenebilir yapı (PyInstaller).

## Kurulum
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

## EXE Paketleme
```bash
pyinstaller --noconfirm --windowed --onefile --name VELINDesk app.py
```

## Not
Bu sürüm bir **MVP/prototip** sürümüdür. Dokusal görseller, fontlar ve animasyonlar isteğe göre daha da gerçekçi hale getirilebilir.
