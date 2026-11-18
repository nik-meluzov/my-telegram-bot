import io
from fastapi import FastAPI, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
import pdfplumber
import openpyxl

app = FastAPI()

@app.get("/")
async def root():
    html_content = """
    <html>
    <head><title>Конвертер PDF</title></head>
    <body style="background:#1e1e1e; color:white; font-family:sans-serif;
                 display:flex; justify-content:center; align-items:center;
                 height:100vh; margin:0;">
        <div style="background:#2d2d2d; padding:40px; border-radius:10px; text-align:center;">
            <h1>Конвертер PDF → Excel</h1>
            <form action="/upload" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".pdf" required>
                

                <button style="background:#007acc; color:white; padding:10px 20px;
                               border:none; border-radius:5px; cursor:pointer;">
                    Загрузить файл
                </button>
            </form>
        </div>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

@app.post("/upload")
async def upload(file: UploadFile):
    content = await file.read()

    if not file.filename.lower().endswith(".pdf"):
        return HTMLResponse(
            content="<h2 style='color:red'>Это не PDF! Выбери файл с расширением .pdf</h2><a href='/' style='color:#007acc'>← Назад</a>"
        )

    try:
        pdf = pdfplumber.open(io.BytesIO(content))
        wb = openpyxl.Workbook()
        first = True
        found_tables = False

        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                found_tables = True
                if first:
                    ws = wb.active
                    ws.title = "Таблица 1"
                    first = False
                else:
                    ws = wb.create_sheet(title=f"Таблица {len(wb.sheetnames) + 1}")
                for row in table:
                    ws.append(row)

        pdf.close()

        if not found_tables:
            return HTMLResponse(
                content="<h2 style='color:orange'>В этом PDF не нашлось таблиц (возможно, это скан или просто текст).</h2><a href='/' style='color:#007acc'>← Назад</a>"
            )

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        xlsx_name = file.filename.rsplit(".", 1)[0] + ".xlsx"
        return StreamingResponse(
            output,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f'attachment; filename="{xlsx_name}"'},
        )
    except Exception as e:
        return HTMLResponse(
            content=f"<h2 style='color:red'>Ошибка: {e}</h2><a href='/' style='color:#007acc'>← Назад</a>"
        )