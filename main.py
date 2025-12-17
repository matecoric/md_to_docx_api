import tempfile
from pathlib import Path
from zipfile import ZipFile
from fastapi import FastAPI, File, UploadFile, Response
import pypandoc  # type: ignore

app = FastAPI()


TEMPLATE_PATH = Path(__file__).parent / "manual_template_styles.docx"


@app.post("/md-to-docx")
async def zip_markdown_to_docx(file: UploadFile = File(..., media_type="application/octet-stream")):
    # Create a temporary working directory
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        # Save uploaded ZIP
        zip_path = tmpdir_path / file.filename
        with open(zip_path, "wb") as f:
            f.write(await file.read())

        # Extract ZIP contents
        with ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(tmpdir_path)

        # Assume there is a single markdown file inside
        md_files = list(tmpdir_path.glob("*.md"))
        if not md_files:
            return {"error": "No markdown file found in the ZIP."}

        md_path = md_files[0]
        output_path = tmpdir_path / "output.docx"

        # Convert Markdown to DOCX
        pypandoc.convert_file(
            str(md_path),
            to="docx",
            format="md",
            outputfile=str(output_path),
            extra_args=[f"--reference-doc={TEMPLATE_PATH}"],
        )

        # Return DOCX file as response
        return Response(
            content=output_path.read_bytes(),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=output.docx"},
        )


@app.get("/health")
def health():
    return {"status": "ok", "pandoc_path": pypandoc.get_pandoc_path()}
