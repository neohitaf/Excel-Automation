from pathlib import Path

from parser import parse_tablo
from templates import FORM3_TEMPLATE
from calc import puantaj_oku, esle, esle_ters, hesapla
from writer import banka_listesi_yaz, bordro_yaz, muhtasar_yaz


def calistir(form3_yolu, puantaj_yolu, cikti_klasoru):
    """
    Form-3 ve aylık puantaj dosyasını işleyerek
    banka listesi, bordro ve muhtasar çıktıları oluşturur.
    """
    cikti_klasoru = Path(cikti_klasoru)
    cikti_klasoru.mkdir(parents=True, exist_ok=True)

    uyarilar = []

    parser_kayitlari, form3_uyarilari = parse_tablo(form3_yolu, FORM3_TEMPLATE)
    uyarilar.extend(form3_uyarilari)

    if not parser_kayitlari:
        raise ValueError("Form-3 dosyasında işlenebilir kayıt bulunamadı.")

    kayitlar, ay = puantaj_oku(puantaj_yolu)

    if not kayitlar:
        raise ValueError("Puantaj dosyasında işlenebilir kayıt bulunamadı.")

    if not ay:
        raise ValueError("Puantaj dosyasından ay bilgisi okunamadı.")

    banka_sonuclari = []
    for parser_kaydi in parser_kayitlari:
        puantaj_kaydi, _ = esle_ters(kayitlar, parser_kaydi)
        birim = puantaj_kaydi.get("birim", "") if puantaj_kaydi else ""
        zenginlesmis = {**parser_kaydi, "birim": birim}
        banka_sonuclari.append(hesapla(zenginlesmis))

    bordro_sonuclari = []
    eslesen_kayit_sayisi = 0
    eslesmeyen_kayit_sayisi = 0
    for puantaj_kaydi in kayitlar:
        parser_kaydi, uyari = esle(parser_kayitlari, puantaj_kaydi)
        if uyari:
            uyarilar.append(uyari)
            parser_kaydi = {}
            eslesmeyen_kayit_sayisi += 1
        else:
            eslesen_kayit_sayisi += 1
        birlesik = {**parser_kaydi, **puantaj_kaydi}
        bordro_sonuclari.append(hesapla(birlesik))

    banka_yolu = cikti_klasoru / "banka_listesi.xlsx"
    bordro_yolu = cikti_klasoru / "bordro.xlsx"
    muhtasar_yolu = cikti_klasoru / "muhtasar.xlsx"

    banka_listesi_yaz(banka_sonuclari, ay, banka_yolu)
    bordro_yaz(bordro_sonuclari, ay, bordro_yolu)
    muhtasar_yaz(bordro_sonuclari, ay, muhtasar_yolu)

    return {
        "ay": ay,
        "uyarilar": uyarilar,
        "istatistikler": {
            "form3_kayit_sayisi": len(parser_kayitlari),
            "puantaj_kayit_sayisi": len(kayitlar),
            "eslesen_kayit_sayisi": eslesen_kayit_sayisi,
            "eslesmeyen_kayit_sayisi": eslesmeyen_kayit_sayisi,
            "banka_kayit_sayisi": len(banka_sonuclari),
            "bordro_kayit_sayisi": len(bordro_sonuclari),
        },
        "cikti_dosyalari": [str(banka_yolu), str(bordro_yolu), str(muhtasar_yolu)],
    }


if __name__ == "__main__":
    try:
        sonuc = calistir(
            form3_yolu="veri/Form-3.docx",
            puantaj_yolu="veri/Puantaj_ornek.xlsx",
            cikti_klasoru="cikti",
        )
        print(f"İşlenen ay: {sonuc['ay']}")
        print("Eşleşen kayıt:", sonuc["istatistikler"]["eslesen_kayit_sayisi"])
        print(f"Toplam {len(sonuc['uyarilar'])} uyarı:")
        for uyari in sonuc["uyarilar"]:
            print("  -", uyari)
        print("\nÜretilen dosyalar:")
        for dosya in sonuc["cikti_dosyalari"]:
            print("  -", dosya)
    except ValueError as exc:
        print("İşlem tamamlanamadı:", exc)