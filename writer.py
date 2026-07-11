import openpyxl
from openpyxl.styles import Font
from templates import BORDRO_SABLON_YOLU
from templates import BORDRO_SABLON_YOLU, BANKA_LISTESI_SABLON_YOLU


def gss_ayir(kayitlar):
    evet_list = []
    hayir_list = []
    for k in kayitlar:
        if k.get("gss") == "EVET":   
            evet_list.append(k)
        else:
            hayir_list.append(k)
    return evet_list, hayir_list

from templates import BORDRO_SABLON_YOLU, BANKA_LISTESI_SABLON_YOLU

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

    wb = openpyxl.load_workbook(BANKA_LISTESI_SABLON_YOLU)
    ws = wb.active

    # Dönem satırını güncelle
    ws.cell(row=5, column=4, value=donem)

    # Veri satırları — 9. satırdan başla
    satir_no = 9
    for k in sirali:
        ws.cell(row=satir_no, column=2, value=k.get("ad_soyad", ""))
        ws.cell(row=satir_no, column=3, value=k.get("tc", ""))
        ws.cell(row=satir_no, column=4, value=k.get("birim", ""))
        ws.cell(row=satir_no, column=5, value=k.get("iban", ""))
        satir_no += 1

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

    wb = openpyxl.load_workbook(BORDRO_SABLON_YOLU)
    ws = wb.active

    ws.cell(row=2, column=1, value=baslik_tarihi)

    # 22 kodlu — satır 6'dan başla
    satir_no = 6
    for k in evet_list:
        ws.cell(row=satir_no, column=2, value=k.get("ad_soyad", ""))
        ws.cell(row=satir_no, column=3, value=k.get("tc", ""))
        ws.cell(row=satir_no, column=4, value=k.get("birim", ""))
        ws.cell(row=satir_no, column=5, value=1101)
        ws.cell(row=satir_no, column=6, value=k.get("prim_gunu", 0))
        satir_no += 1

    # 43 kodlu — sabit satır 84'ten başla
    satir_no = 84
    for k in hayir_list:
        ws.cell(row=satir_no, column=2, value=k.get("ad_soyad", ""))
        ws.cell(row=satir_no, column=3, value=k.get("tc", ""))
        ws.cell(row=satir_no, column=4, value=k.get("birim", ""))
        ws.cell(row=satir_no, column=5, value=1101)
        ws.cell(row=satir_no, column=6, value=k.get("prim_gunu", 0))
        satir_no += 1

    wb.save(cikti_yolu)
    print(f"Kaydedildi: {cikti_yolu}")
