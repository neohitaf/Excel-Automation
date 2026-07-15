from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from main import form3_ayristir, coklu_puantaj_isle

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="web_templates")


@app.get("/", response_class=HTMLResponse)
def anasayfa(request: Request):
    return templates.TemplateResponse(request=request, name="index.html", context={})


@app.post("/isle", response_class=HTMLResponse)
async def dosyalari_isle(
    request: Request,
    form3: UploadFile = File(...),
    puantaj_listesi: list[UploadFile] = File(...),
    beklenen_ay: str = Form(...),
):
    is_kimligi = str(uuid4())
    is_klasoru = Path("instance/jobs") / is_kimligi
    girdi_klasoru = is_klasoru / "girdi"
    cikti_klasoru = is_klasoru / "cikti"

    girdi_klasoru.mkdir(parents=True, exist_ok=True)

    form3_yolu = girdi_klasoru / "Form-3.docx"
    form3_yolu.write_bytes(await form3.read())

    puantaj_yollari = []
    for i, puantaj_dosyasi in enumerate(puantaj_listesi):
        puantaj_yolu = girdi_klasoru / f"Puantaj_{i}.xlsx"
        puantaj_yolu.write_bytes(await puantaj_dosyasi.read())
        puantaj_yollari.append(str(puantaj_yolu))

    try:
        parser_kayitlari, form3_uyarilari = form3_ayristir(str(form3_yolu))
        sonuc = coklu_puantaj_isle(
            parser_kayitlari,
            form3_uyarilari,
            puantaj_yollari,
            str(cikti_klasoru),
            beklenen_ay=beklenen_ay,
        )
        return templates.TemplateResponse(
            request=request,
            name="sonuc.html",
            context={"sonuc": sonuc, "is_kimligi": is_kimligi, "hata": None},
        )
    except ValueError as hata:
        return templates.TemplateResponse(
            request=request,
            name="sonuc.html",
            context={"sonuc": None, "is_kimligi": None, "hata": str(hata)},
        )


@app.get("/indir/{is_kimligi}/{dosya_adi}")
def dosya_indir(is_kimligi: str, dosya_adi: str):
    dosya_yolu = Path("instance/jobs") / is_kimligi / "cikti" / dosya_adi
    return FileResponse(path=dosya_yolu, filename=dosya_adi)