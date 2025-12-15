import tempfile
from pathlib import Path

import pypandoc  # type: ignore
from fastapi import Body, FastAPI, Response


app = FastAPI()

TEMPLATE_PATH = Path(__file__).parent / "manual_template_styles.docx"


@app.post("/md-to-docx")
def markdown_to_docx(markdown_string: str = Body(..., media_type="text/plain")):
    with tempfile.TemporaryDirectory() as tmpdir:
        output_path = Path(tmpdir) / "output.docx"

        pypandoc.convert_text(
            source=markdown_string,
            to="docx",
            format="md",
            outputfile=str(output_path),
            extra_args=[f"--reference-doc={TEMPLATE_PATH}"],
        )

        return Response(
            content=output_path.read_bytes(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=output.docx"},
        )


@app.get("/health")
def health():
    return {"status": "ok", "pandoc_path": pypandoc.get_pandoc_path()}
