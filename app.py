import os
import base64
import tempfile
import shutil
import importlib
from typing import Optional, List

from fastapi import FastAPI, HTTPException
from fastapi.responses import JSONResponse
from pydantic import BaseModel

app = FastAPI()


class ImageItem(BaseModel):
    # Ex.: "Foto_ambiente_Images/cc84ffd0.Foto.190306.jpg"
    path: str
    # conteúdo do arquivo em base64
    base64: str


class Payload(BaseModel):
    id_vistoria: str
    excel_base64: str
    template_base64: Optional[str] = None
    images: Optional[List[ImageItem]] = None


@app.post("/generate")
def generate(p: Payload):
    work = tempfile.mkdtemp(prefix="laudo_")
    try:
        # Workspace temporário (Render)
        os.environ["LAUDO_BASE_DIR"] = work

        # 1) Salvar o Excel
        excel_path = os.path.join(work, "Vistoria.xlsx")
        with open(excel_path, "wb") as f:
            f.write(base64.b64decode(p.excel_base64))

        # 2) Salvar o template (preferir o enviado; senão usar o do repositório)
        dst_template = os.path.join(work, "Modelo_Vistoria.docx")
        if p.template_base64:
            with open(dst_template, "wb") as f:
                f.write(base64.b64decode(p.template_base64))
        else:
            template_src = os.path.join(os.path.dirname(__file__), "Modelo_Vistoria.docx")
            shutil.copy(template_src, dst_template)

        # 3) Salvar imagens recebidas respeitando a estrutura de pastas
        if p.images:
            for it in p.images:
                rel = (it.path or "").lstrip("/").replace("\\", "/")
                abs_path = os.path.join(work, rel)
                os.makedirs(os.path.dirname(abs_path), exist_ok=True)
                with open(abs_path, "wb") as f:
                    f.write(base64.b64decode(it.base64))

        # 4) Importar gerar_laudo APÓS setar LAUDO_BASE_DIR
        import gerar_laudo as gl
        importlib.reload(gl)

        # 5) Gerar o laudo
        out_path = gl.gerar_laudo(p.id_vistoria)

        # 6) Ler o DOCX gerado e devolver base64
        if not os.path.isfile(out_path):
            raise HTTPException(500, "Laudo não foi gerado (arquivo .docx não encontrado).")

        filename = os.path.basename(out_path)
        with open(out_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode("utf-8")

        return JSONResponse({"filename": filename, "docx_base64": docx_b64})

    except Exception as e:
        import traceback
        print("=== ERRO NO /generate ===")
        print(traceback.format_exc())
        raise HTTPException(500, f"Erro gerando laudo: {e}")

    finally:
        shutil.rmtree(work, ignore_errors=True)
