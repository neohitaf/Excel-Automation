import openpyxl

GUNLUK_UCRET = 1101
SGK_AYI = 30


def puantaj_oku(dosya_yolu):
    """
    Puantaj Excel'ini okur.
    Döner: [{"ad_soyad": "Buse ZEYBEK", "toplam_saat": 32}, ...]
    """
    wb = openpyxl.load_workbook(dosya_yolu, data_only=True)
    ws = wb.active

    kayitlar = []

    # Satır 14'ten başla (13 başlık), ad boşsa atla
    for row in ws.iter_rows(min_row=14):
        ad_soyad = row[1].value   # sütun B → index 1
        toplam   = row[34].value  # sütun AI → index 34
        gorev     = row[2].value   # sütun C → index 2
        if not ad_soyad:
            continue
        if not gorev:
            continue
        if str(gorev).strip() != "Öğrenci":
            continue

        kayitlar.append({
            "ad_soyad":    str(ad_soyad).strip(),
            "toplam_saat": int(toplam) if toplam else 0,
        })
        

    return kayitlar


def hesapla(kayit):
    """
    Tek öğrenci için mali hesapları yapar.
    Girdi:  {"ad_soyad": ..., "toplam_saat": 32, "gss": "EVET"}
    Döner:  aynı dict + hesaplanan alanlar eklenerek
    """
    saat        = kayit["toplam_saat"]
    prim_gunu   = saat // 8          # tam gün, küsurat yok
    ucret       = prim_gunu * GUNLUK_UCRET
    eksik_gun   = SGK_AYI - prim_gunu
    belge_turu  = 22 if kayit.get("gss") == "EVET" else 43

    return {
        **kayit,                     # mevcut alanları koru
        "prim_gunu":   prim_gunu,
        "ucret":       ucret,
        "eksik_gun":   eksik_gun,
        "belge_turu":  belge_turu,
    }


if __name__ == "__main__":
    kayitlar = puantaj_oku("veri/Puantaj_ornek.xlsx")
    for k in kayitlar:
        k["gss"] = "EVET"
        sonuc = hesapla(k)
        print(sonuc)
        print()