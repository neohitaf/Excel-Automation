# Her tablo tipi için şablon tanımları.
# Zorunlu: olmadan işlenemez, hata fırlatır.
# Opsiyonel: yoksa uyarı verir, devam eder.

FORM3_TEMPLATE = {
    "zorunlu": {
        "ad_soyad":      ["ADI SOYADI", "AD SOYAD", "ADINIZ"],
        "tc":            ["KİMLİK", "T.C", "TC NO"],
        "gss":           ["G.S.S", "GSS", "SAĞLIK"],
        "iban":          ["IBAN", "BAN NO", "HESAP"],
        "banka":         ["BANKA"],
    },
    "opsiyonel": {
        "sira":          ["SIRA", "S.NO", "NO"],
        "telefon":       ["TELEFON", "CEP", "GSM"],
        "calisma_sekli": ["ÇALIŞMA"],
        "bolum":         ["BÖLÜM", "PROGRAM", "ANABİLİM"],
    }
}

# İleride başka bir tablo eklemek istersen:
# BORDRO_TEMPLATE = {
#     "zorunlu": { "ad_soyad": [...], "tc": [...], ... },
#     "opsiyonel": { ... }
# }