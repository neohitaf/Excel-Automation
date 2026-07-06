import openpyxl
from openpyxl.styles import Font

def banka_listesi_yaz(kayitlar, cikti_yolu):
    evet_list = []
    hayir_list = []

    for k in kayitlar:
        if k["gss"] == "EVET":
            evet_list.append(k)
        else:
            hayir_list.append(k)

    sirali = evet_list + hayir_list

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Banka Listesi"

    # Başlık satırı
    basliklar = ["SIRA", "AD SOYAD", "TC KİMLİK NO", "BANKA", "IBAN", "NET ÜCRET", "GSS"]
    ws.append(basliklar)

    # Başlıkları kalın yap
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # Veri satırları
    for sira, k in enumerate(sirali, start=1):
        ws.append([
            sira,
            k["ad_soyad"],
            k["tc"],
            k.get("banka", ""),
            k.get("iban", ""),
            k["ucret"],
            k["gss"],
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