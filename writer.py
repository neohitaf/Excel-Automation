import openpyxl
from openpyxl.styles import Font


def gss_ayir(kayitlar):
    evet_list = []
    hayir_list = []
    for k in kayitlar:
        if k.get("gss") == "EVET":   
            evet_list.append(k)
        else:
            hayir_list.append(k)
    return evet_list, hayir_list



def banka_listesi_yaz(kayitlar, ay, cikti_yolu):
    from datetime import datetime
    import calendar

    AY_NUMARALARI = {"OCAK": 1, "ŞUBAT": 2, "MART": 3, "NİSAN": 4,
                     "MAYIS": 5, "HAZİRAN": 6, "TEMMUZ": 7, "AĞUSTOS": 8,
                     "EYLÜL": 9, "EKİM": 10, "KASIM": 11, "ARALIK": 12}
    yil = datetime.now().year
    ay_no = AY_NUMARALARI.get(ay, 1)
    son_gun = calendar.monthrange(yil, ay_no)[1]
    donem = f"01 {ay}- {son_gun} {ay} {yil}"

    evet_list, hayir_list = gss_ayir(kayitlar)
    sirali = evet_list + hayir_list

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Banka Listesi"

    # Üst bilgi bloğu — sabit
    ws.append(["BANKA LİSTESİ"])
    ws.append(["Muhasebe Biriminin Adı-Kodu", "", "", "Strateji Geliştirme Daire Başkanlığı", ""])
    ws.append(["Harcama Biriminin Adı- Kodu", "", "", "Sağlık Kültür ve Spor Daire Başkanlığı", "", "907"])
    ws.append(["Banka Adı", "", "", "Yapı Kredi Bankası", ""])
    ws.append(["Aylığın Ait Olduğu Yıl-Dönem", "", "", donem, ""])   # ← sadece bu değişiyor
    ws.append(["BANKA HESAP NO", "ALACAKLILARIN"])

    # Sütun başlıkları
    ws.append(["Sıra\nNO", "ADI SOYADI", "T.C.\nKİMLİK NO", "ÇALIŞTIĞI BİRİM",
               "BANKA \nHESAP NOSU/IBAN", "TOPLAM \nELE GEÇEN"])

    for cell in ws[7]:
        cell.font = Font(bold=True)

    for sira, k in enumerate(sirali, start=1):
        ws.append([
            sira,
            k.get("ad_soyad", ""),
            k.get("tc", ""),
            k.get("birim", ""),
            k.get("iban", ""),
            k.get("ucret", ""),
        ])

    for col in ws.columns:
        max_uzunluk = 0
        sutun_harfi = col[0].column_letter
        for cell in col:
            if cell.value:
                max_uzunluk = max(max_uzunluk, len(str(cell.value)))
        ws.column_dimensions[sutun_harfi].width = max_uzunluk + 4

    wb.save(cikti_yolu)
    print(f"Kaydedildi: {cikti_yolu}")



def bordro_yaz(kayitlar, ay, cikti_yolu):
    from datetime import datetime
    import calendar

    AY_NUMARALARI = {"OCAK": 1, "ŞUBAT": 2, "MART": 3, "NİSAN": 4,
                     "MAYIS": 5, "HAZİRAN": 6, "TEMMUZ": 7, "AĞUSTOS": 8,
                     "EYLÜL": 9, "EKİM": 10, "KASIM": 11, "ARALIK": 12}
    yil = datetime.now().year
    ay_no = AY_NUMARALARI.get(ay, 1)
    son_gun = calendar.monthrange(yil, ay_no)[1]
    baslik_tarihi = f"1 {ay} {yil} - {son_gun} {ay} {yil}"

    evet_list, hayir_list = gss_ayir(kayitlar)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Bordro"

    # Başlık satırları
    ws.append(["KISMİ ZAMANLI MAAŞ BORDROSU"])
    ws.append([baslik_tarihi])
    ws.append([])
    ws.append(["Sıra No", "Adı Soyadı", "T.C. Kimlik No", "Fakülte", "Günlüğü",
               "Toplam Gün", "Prime Esas Kazanç", "Gelir Vergisi Matrahı",
               "SGK %1", "SGK %5", "Genel Toplam", "SGK %1", "SGK %5",
               "İcra Kesintisi", "Kesintiler Toplamı", "Net Ücret"])

    # 22 kodlu bölüm
    ws.append(["22 - KODLU ÖĞRENCİLER"])
    for sira, k in enumerate(evet_list, start=1):
        ws.append([
            sira,
            k.get("ad_soyad", ""),
            k.get("tc", ""),
            k.get("birim", ""),
            1101,
            k.get("prim_gunu", 0),
            "", "", "", "", "", "", "", "", "", ""
        ])
    evet_toplam = sum(k.get("ucret", 0) for k in evet_list)
    ws.append(["TOPLAM", "", "", "", "", "", "", "", "", "", "", "", "", "", "", evet_toplam])

    # 43 kodlu bölüm
    ws.append([])
    ws.append(["43 - KODLU ÖĞRENCİLER"])
    for sira, k in enumerate(hayir_list, start=len(evet_list) + 1):
        ws.append([
            sira,
            k.get("ad_soyad", ""),
            k.get("tc", ""),
            k.get("birim", ""),
            1101,
            k.get("prim_gunu", 0),
            "", "", "", "", "", "", "", "", "", ""
        ])
    hayir_toplam = sum(k.get("ucret", 0) for k in hayir_list)
    ws.append(["TOPLAM", "", "", "", "", "", "", "", "", "", "", "", "", "", "", hayir_toplam])

    # Genel toplam
    ws.append([])
    genel_toplam = evet_toplam + hayir_toplam
    ws.append(["GENEL TOPLAM", "", "", "", "", "", "", "", "", "", "", "", "", "", "", genel_toplam])

    # Sütun genişliği + kaydet
    for col in ws.columns:
        max_uzunluk = 0
        sutun_harfi = col[0].column_letter
        for cell in col:
            if cell.value:
                max_uzunluk = max(max_uzunluk, len(str(cell.value)))
        ws.column_dimensions[sutun_harfi].width = max_uzunluk + 4

    wb.save(cikti_yolu)
    print(f"Kaydedildi: {cikti_yolu}")
