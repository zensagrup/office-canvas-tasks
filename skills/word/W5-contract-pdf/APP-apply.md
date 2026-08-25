# Uygulama — W5 Sözleşme → Ayrı PDF

Yöntem A — Word yerleşik (tek belge):
1. Mail Merge ile {FirstName}{Amount} yer tutucuları.
2. **Finish & Merge** → **Edit Individual Documents** → hepsi tek belgede.
3. Sonra **Böl** (VBA/kenar boşlukları) veya her sayfayı ayrı kaydet.

Yöntem B — Aktif Merge / Plumsail (önerilen):
- Excel'i yükle, Word .docx şablonunu yükle ({FirstName} tarzı), alan eşlemei önizle.
- **Generate All** → her kayıt ayrı PDF → ZIP.

> Word yerleşik tek dosya üretir; "ayrı dosya" istersen yöntem B (ActiveMerge/Plumsail) veya power-int.
