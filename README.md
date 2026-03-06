# Multimetre Mastar Excel Üretici

Bu proje, multimetre marka/model bilgisi ve ölçüm aralıklarına göre kalibrasyon mastarı Excel dosyası (`.xlsx`) üretir.

## Kullanım

```bash
python3 generate_mastar.py \
  --brand "Fluke" \
  --model "87V" \
  --ranges "0.2,2,20,200,1000" \
  --output "fluke_87v_mastar.xlsx"
```

## Üretilen Excel İçeriği

- Marka / model bilgisi
- Ölçüm aralığı
- %10, %50, %90 test noktaları
- Kalibrasyon için tablo düzeni

> `--ranges` içinde verilen her aralık için `%10`, `%50`, `%90` değerleri otomatik hesaplanır.
