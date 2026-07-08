# Poligon PPTX Daily Report

CS Team Daily Chat Stats sunumunu Comm100 API'sinden otomatik dolduran, mail eki olarak gönderen script.
`erhanbs70/poligon-chat-analyzer` reposu altına bu klasörü (`pptx-report/` gibi) ekleyip kullanabilirsin —
mevcut GAS scriptlerinden (v9 Daily Report, VIP Daily Report) tamamen bağımsız çalışır, birbirlerini
etkilemezler.

## Dosyalar

| Dosya | Görev |
|---|---|
| `fetch_metrics.py` | Comm100 API'den chat + reporting verisini çeker (v9 GAS script'inin Python portu) |
| `template.pptx` | `{{TOKEN}}` placeholder'lı şablon (1 kere hazırlandı, `make_template.py` ile üretildi) |
| `build_pptx.py` | Şablonu metriklerle doldurur, grafikleri üretir (matplotlib), pptx'i kaydeder |
| `main.py` | Orkestrasyon: fetch → build → Gmail ile gönder |
| `make_template.py` | Şablonu orijinal pptx'ten üreten script (tasarım değişmedikçe tekrar çalıştırmana gerek yok) |
| `test_mock_build.py` | Gerçek API'ye dokunmadan, sahte veriyle pipeline'ı test etmek için |

## Kurulum

### 1. GitHub Secrets ekle (repo → Settings → Secrets and variables → Actions)

| Secret | Açıklama |
|---|---|
| `MR_SITE_ID` | Comm100 site ID (`90005373`) |
| `MR_API_KEY` | Comm100 API key |
| `MR_EMAIL` | Comm100 API email (`erhan@premilogic.com`) |
| `MR_CS_DEPT` | Customer Support departman GUID |
| `GMAIL_ADDRESS` | Mail gönderecek Gmail adresi |
| `GMAIL_APP_PASSWORD` | Bu adresin **App Password**'ü (normal şifre değil — Google Hesap → Güvenlik → 2 Adımlı Doğrulama → Uygulama Şifreleri'nden üretilir, 2FA açık olmalı) |
| `MAIL_TO` | Alıcılar, virgülle ayrılmış (`erhanbs70@gmail.com,tahircs252@gmail.com,...`) |

### 2. Klasörü repoya ekle

Bu klasörün tamamını (workflow dahil) `erhanbs70/poligon-chat-analyzer` reposuna kopyala. Workflow dosyasındaki
`working-directory: pptx-report` satırını, klasörü hangi isimle/yola koyduysan ona göre güncelle.

### 3. Test et (lokal, opsiyonel ama önerilir)

```bash
pip install -r requirements.txt
python3 test_mock_build.py          # sahte veriyle, API'ye dokunmadan hızlı test
python3 fetch_metrics.py 2026-07-06 # gerçek API'den metrik çek, JSON olarak yazdır (env vars gerekli)
python3 main.py 2026-07-06          # tam akış: fetch + build + mail gönder
```

### 4. Workflow'u tetikle

Otomatik: her gün 07:35 Sofia'da (`cron: "35 4 * * *"`, UTC bazlı, yaz saatinde 3 saat fark).
Manuel: GitHub → Actions → "Daily PPTX Report" → "Run workflow" (belirli bir tarih girebilirsin).

## Şu an hazır olan 7 slayt

Slide 1 (tarih), 2 (Missed Chats genel), 3 (Response/Duration), 4 (Chatbot genel),
5-6-7 (brand bazlı Missed + Chatbot). Slide 9 (email/call istatistikleri) bu sürümde **çıkarıldı**
çünkü o veri Comm100 dışında bir kaynaktan geliyor (muhtemelen Zendesk + call panel Sheets) — kaynak
netleşince eklenir.

## Durum: 7 slaytın tamamı artık tam otomatik, hiç "N/A" yok 🎉

Tüm veri gerçek Comm100 reporting API'sinden geliyor, hiçbir tahmin/yaklaşım yok:

- Slide 2, 5, 6, 7 — saat bazlı Missed/Served/Success gerçek (`timeDisplayType: "hour"`), en kötü
  3 saat gerçek missed sayısına göre seçiliyor, grafikler served+missed stacked bar + acceptance line
- Slide 3 — Avg. First Response Time dahil tüm efficiency metrikleri gerçek (`AgentFirstResponseTime`)
- Slide 4 — Action usage / Action per chat gerçek (`AIReplyRecord` cube entity, `UsedAIReplies`,
  payda = `botToAgent`)
- Slide 8 — marka bazlı 3 donut gerçek (Wrap-up survey reporting API, `campaignId + categoryOptionName`)

Slide 9 (email/call istatistikleri) hâlâ bu sürümde yok — o veri Comm100 dışında bir kaynaktan
geliyor (muhtemelen Zendesk + call panel Sheets), kaynak netleşince eklenir.

