# Uygulama Araçları (apps/)

Bu klasördeki repo'lar **ücretsiz / açık kaynak** araçlardır ve `skills/` altındaki
senaryolarla (Excel / Word / Canvas) birebir örtüşür. Hermes `git clone --depth 1` ile indirdi.

## Excel
| Repo | Lisans | Karşılar |
|---|---|---|
| `excel/xlwings` | BSD (açık kaynak kısmı) | E1-E6: Power Query yerine Python, VBA yerine Python UDF, raporlama |
| `excel/excel-automation-toolkit` | MIT | E4/E5: pivot, VLOOKUP/XLOOKUP, hesaplamalı sütun tek fonksiyonla |

## Word
| Repo | Lisans | Karşılar |
|---|---|---|
| `word/docx-mailmerge` | MIT | W1-W5: Office yüklü olmadan docx mail merge (mektup, sertifika, sözleşme→PDF) |

## Canvas
| Repo | Lisans | Karşılar |
|---|---|---|
| `canvas/canva-connect-api-starter-kit` | Canva (resmi, ücretsiz) | C1-C4: Canva Connect API ile marka uyumlu tasarım, Magic Switch, autofill |

## Kurulum
```bash
# Python araçları
pip install xlwings openpyxl pandas
pip install docx-mailmerge2
# excel-automation-toolkit
cd apps/excel/excel-automation-toolkit && pip install -r requirements.txt
# Canva starter kit (Node)
cd apps/canvas/canva-connect-api-starter-kit && npm install
```

## Not
- `xlwings` PRO klasörü kapalı kaynaktır; açık kaynak kısmı (BSD) yeterli.
- Canva API ücretsiz planda rate limit'li (örn. generate-design 20 req/min); ücretli özellikler ayrı.
