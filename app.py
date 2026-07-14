from pathlib import Path
from uuid import uuid4

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from main import calistir

app = FastAPI()

app.mount("/static", StaticFiles(directory="static"), name="static")

templates = Jinja2Templates(directory="web_templates")


@app.get("/", response_class=HTMLResponse)
def anasayfa(request: Request):
    """Dosya yükleme formunu gösterir."""
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={},
    )


@app.post("/isle", response_class=HTMLResponse)
async def dosyalari_isle(
    request: Request,
    form3: UploadFile = File(...),
    puantaj: UploadFile = File(...),
):
    """Yüklenen dosyaları alır, calistir() ile işler, sonucu gösterir."""

    is_kimligi = str(uuid4())
    is_klasoru = Path("instance/jobs") / is_kimligi
    girdi_klasoru = is_klasoru / "girdi"
    cikti_klasoru = is_klasoru / "cikti"

    girdi_klasoru.mkdir(parents=True, exist_ok=True)
    cikti_klasoru.mkdir(parents=True, exist_ok=True)

    form3_yolu = girdi_klasoru / "Form-3.docx"
    puantaj_yolu = girdi_klasoru / "Puantaj.xlsx"

    form3_yolu.write_bytes(await form3.read())
    puantaj_yolu.write_bytes(await puantaj.read())

    try:
        sonuc = calistir(
            form3_yolu=str(form3_yolu),
            puantaj_yolu=str(puantaj_yolu),
            cikti_klasoru=str(cikti_klasoru),
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
    """Üretilen Excel dosyasını indirir."""
    dosya_yolu = Path("instance/jobs") / is_kimligi / "cikti" / dosya_adi
    return FileResponse(path=dosya_yolu, filename=dosya_adi)