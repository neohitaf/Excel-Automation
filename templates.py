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

BORDRO_SABLON_YOLU = "templates/bordro_sablon.xlsx"
BANKA_LISTESI_SABLON_YOLU = "templates/banka_listesi_sablon.xlsx"
MUHTASAR_SABLON_YOLU = "templates/muhtasar_sablon.xlsx"

# MUHTASAR sabit değerleri — her öğrenci için aynı
MUHTASAR_SABIT = {
    "belgenin_mahiyeti":     "A",
    "kanun_no":              "00000",
    "yeni_unite_kodu":       "01",
    "eski_unite_kodu":       "01",
    "isyeri_sira_no":        "1143127",
    "il_kodu":               "042",
    "alt_isveren_no":        "000",
    "eksik_gun_nedeni":      "7",
    "meslek_kodu":           "9901.02",
    "tahakkuk_nedeni":       "A",
    "gelir_vergisi_muaf":    2,
    "gv_engellilik_orani":   0,
    "gv_kesintisi":          0,
    "damga_vergisi_kesintisi": 0,
}


