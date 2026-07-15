from pathlib import Path

from parser import parse_tablo
from templates import FORM3_TEMPLATE
from calc import puantaj_oku, esle, esle_ters, hesapla
from writer import banka_listesi_yaz, bordro_yaz, muhtasar_yaz


def form3_ayristir(form3_yolu):
    parser_kayitlari, form3_uyarilari = parse_tablo(form3_yolu, FORM3_TEMPLATE)
    if not parser_kayitlari:
        raise ValueError("Form-3 dosyasında işlenebilir kayıt bulunamadı.")
    return parser_kayitlari, form3_uyarilari


def _kayitlari_isle_ve_yaz(parser_kayitlari, form3_uyarilari, kayitlar, ay, cikti_klasoru):
    """
    Ortak mantık: eşleştirme + hesaplama + 3 Excel dosyasını yazma.
    Hem tek Puantaj hem birleştirilmiş çoklu Puantaj kayıtları için kullanılır.
    """
    cikti_klasoru = Path(cikti_klasoru)
    cikti_klasoru.mkdir(parents=True, exist_ok=True)

    uyarilar = list(form3_uyarilari)

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


def puantaj_isle(parser_kayitlari, form3_uyarilari, puantaj_yolu, cikti_klasoru, beklenen_ay=None):
    """Tek bir Puantaj dosyasını işler — CLI ve tekli kullanım için."""
    kayitlar, ay = puantaj_oku(puantaj_yolu)

    if not kayitlar:
        raise ValueError("Puantaj dosyasında işlenebilir kayıt bulunamadı.")
    if not ay:
        raise ValueError("Puantaj dosyasından ay bilgisi okunamadı.")
    if beklenen_ay and ay != beklenen_ay:
        raise ValueError(
            f"Bu dosyadan okunan ay '{ay}', ancak seçilen/beklenen ay '{beklenen_ay}'."
        )

    return _kayitlari_isle_ve_yaz(parser_kayitlari, form3_uyarilari, kayitlar, ay, cikti_klasoru)


def coklu_puantaj_isle(parser_kayitlari, form3_uyarilari, puantaj_yollari, cikti_klasoru, beklenen_ay):
    """
    Birden fazla Puantaj dosyasını okuyup öğrencilerini BİRLEŞTİRİR,
    tek seferde eşleştirip TEK bir 3-dosyalık çıktı üretir.
    Ay uyuşmayan veya okunamayan dosyalar atlanır, atlanma sebebiyle raporlanır.
    """
    tum_kayitlar = []
    atlanan_dosyalar = []

    for puantaj_yolu in puantaj_yollari:
        kayitlar, ay = puantaj_oku(puantaj_yolu)

        if not kayitlar:
            atlanan_dosyalar.append({
                "dosya_adi": Path(puantaj_yolu).name,
                "hata": "İşlenebilir kayıt bulunamadı.",
            })
            continue

        if not ay:
            atlanan_dosyalar.append({
                "dosya_adi": Path(puantaj_yolu).name,
                "hata": "Ay bilgisi okunamadı.",
            })
            continue

        if ay != beklenen_ay:
            atlanan_dosyalar.append({
                "dosya_adi": Path(puantaj_yolu).name,
                "hata": f"Bu dosyadan okunan ay '{ay}', ancak seçilen ay '{beklenen_ay}'.",
            })
            continue

        tum_kayitlar.extend(kayitlar)

    if not tum_kayitlar:
        raise ValueError(
            "Hiçbir Puantaj dosyasından işlenebilir kayıt elde edilemedi — "
            "tüm dosyalar atlandı, seçilen ay ile uyuşmuyor olabilir."
        )

    sonuc = _kayitlari_isle_ve_yaz(parser_kayitlari, form3_uyarilari, tum_kayitlar, beklenen_ay, cikti_klasoru)
    sonuc["atlanan_dosyalar"] = atlanan_dosyalar
    return sonuc


def calistir(form3_yolu, puantaj_yolu, cikti_klasoru, beklenen_ay=None):
    """Tek dosyalık kullanım için — CLI ve geriye dönük uyumluluk."""
    parser_kayitlari, form3_uyarilari = form3_ayristir(form3_yolu)
    return puantaj_isle(parser_kayitlari, form3_uyarilari, puantaj_yolu, cikti_klasoru, beklenen_ay)


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