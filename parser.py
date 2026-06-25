from docx import Document
from pathlib import Path
from templates import FORM3_TEMPLATE


# ─── Yardımcı (private) fonksiyonlar 
"""" fonksiyonlar, şablonla eşleşen tabloyu bulmak ve başlıkları eşleştirmek için kullanılır. """
def _sutun_bul(alan_adi, anahtar_kelimeler, basliklar):
    for i, baslik in enumerate(basliklar):
        if any(kelime in baslik for kelime in anahtar_kelimeler):
            return i
    return None


def _banka_temizle(banka):
    """'YAPI KREDİ' ve 'YAPIKREDİ' gibi tutarsız yazımları birleştirir."""
    return banka.replace(" ", "").upper()


def _gss_temizle(gss):
    """'Evet', 'evet', 'EVET' → 'EVET' | 'Hayır' → 'HAYIR'"""
    return gss.strip().upper()

def _ad_soyad_ayir(ad_soyad):
    """
    'NURETTİN TALHA UZUN' → {"ad": "NURETTİN TALHA", "soyad": "UZUN"}
    İkiden fazla kelime varsa uyarı döner.
    """
    parcalar = ad_soyad.strip().split()

    if len(parcalar) == 0:
        return {"ad": "", "soyad": ""}, "Ad soyad boş geldi."

    if len(parcalar) == 1:
        return {"ad": parcalar[0], "soyad": ""}, "Tek kelime var, soyad bulunamadı."

    soyad = parcalar[-1]       # son kelime
    ad = " ".join(parcalar[:-1])  # son kelime hariç hepsi

    uyari = None
    if len(parcalar) > 2:
        uyari = f"'{ad_soyad}' → ad: '{ad}' | soyad: '{soyad}' — birleşik soyad olabilir, kontrol edin."

    return {"ad": ad, "soyad": soyad}, uyari


# ─── Ana yardımcı fonksiyonlar 

def sutun_haritasi_olustur(tablo, template):
    basliklar = [
        cell.text.strip().upper().replace("\n", " ")
        for cell in tablo.rows[0].cells
    ]

    konum_haritasi = {}
    uyarilar = []
    eksik_zorunlular = []

    for key, value in template["zorunlu"].items():
        sutun_no = _sutun_bul(key, value, basliklar)
        if sutun_no is None:
            eksik_zorunlular.append(key)
        else:
            konum_haritasi[key] = sutun_no

    if eksik_zorunlular:
        raise ValueError(
            f"Zorunlu başlıklar bulunamadı: {', '.join(eksik_zorunlular)}\n"
            f"Tablodaki başlıklar: {basliklar}"
        )

    for key, value in template.get("opsiyonel", {}).items():
        sutun_no = _sutun_bul(key, value, basliklar)
        if sutun_no is None:
            uyarilar.append(f"Opsiyonel alan bulunamadı: '{key}'")
        else:
            konum_haritasi[key] = sutun_no

    return konum_haritasi, uyarilar


# ─── Dışarıya açık ana fonksiyon

def parse_tablo(dosya_yolu, template):
    dosya = Path(dosya_yolu)
    if not dosya.exists():
        raise FileNotFoundError(f"Dosya bulunamadı: {dosya_yolu}")

    bel = Document(dosya_yolu)

    hedef_tablo = None
    konum_haritasi = None
    tablo_uyarilari = []

    for tablo in bel.tables:
        if len(tablo.rows) == 0:
            continue
        try:
            konum_haritasi, tablo_uyarilari = sutun_haritasi_olustur(tablo, template)
            hedef_tablo = tablo
            break
        except ValueError:
            continue

    if hedef_tablo is None:
        raise ValueError(f"{dosya.name}: Şablona uyan tablo bulunamadı.")

    kayitlar = []
    uyarilar = list(tablo_uyarilari)

    for satir_no, satir in enumerate(hedef_tablo.rows[1:], start=1):
        hucreler = [
            cell.text.strip().replace("\n", " ")
            for cell in satir.cells
        ]

        if "ad_soyad" in konum_haritasi:
            if not hucreler[konum_haritasi["ad_soyad"]]:
                continue

        kayit = {"kaynak_dosya": dosya.name}
        satir_eksikleri = []

        for key, sutun_no in konum_haritasi.items():
            deger = hucreler[sutun_no] if sutun_no < len(hucreler) else ""
            kayit[key] = deger
            if key in template["zorunlu"] and not deger:
                satir_eksikleri.append(key)

        # Temizleme — kayıt oluşturulduktan hemen sonra uygula
        if "banka" in kayit:
            kayit["banka"] = _banka_temizle(kayit["banka"])
        if "gss" in kayit:
            kayit["gss"] = _gss_temizle(kayit["gss"])

        if "ad_soyad" in kayit:
            ayrim, ad_uyari = _ad_soyad_ayir(kayit["ad_soyad"])
            kayit["ad"] = ayrim["ad"]
            kayit["soyad"] = ayrim["soyad"]
            if ad_uyari:
                uyarilar.append(f"{dosya.name} satır {satir_no}: {ad_uyari}")    

        if satir_eksikleri:
            isim = kayit.get("ad_soyad", "?")
            uyarilar.append(
                f"{dosya.name} satır {satir_no} ({isim}): "
                f"eksik zorunlu alan: {', '.join(satir_eksikleri)}"
            )

        kayitlar.append(kayit)

    if not kayitlar:
        uyarilar.append(f"{dosya.name}: Hiç kayıt bulunamadı.")

    return kayitlar, uyarilar


# ─── Test bloğu

if __name__ == "__main__":
    kayitlar, uyarilar = parse_tablo(
        "veri/Form-3 Deneme.docx",
        FORM3_TEMPLATE
    )

    print(f"Bulunan kayıt sayısı: {len(kayitlar)}")
    print()
    for kayit in kayitlar:
        print(kayit)
        print()

    if uyarilar:
        print("--- UYARILAR ---")
        for u in uyarilar:
            print(f"  ⚠  {u}")