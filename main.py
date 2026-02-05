from fastapi import FastAPI, UploadFile, File, Form
from pptx import Presentation
from io import BytesIO
import json

app = FastAPI()

@app.post("/rebuild-ppt")
async def rebuild_ppt(file: UploadFile = File(...), content: str = Form(...)):
    new_data = json.loads(content)
    # Open the original file as a template
    prs = Presentation(BytesIO(await file.read()))

    for i, slide in enumerate(prs.slides):
        if i < len(new_data):
            slide_update = new_data[i]
            for shape in slide.shapes:
                if not shape.has_text_frame: continue
                
                # Surgical replacement in the first run to keep formatting
                if shape == slide.shapes.title:
                    shape.text_frame.paragraphs[0].runs[0].text = slide_update['new_title']
                else:
                    combined_text = "\n".join(slide_update['new_content'])
                    shape.text_frame.paragraphs[0].runs[0].text = combined_text

    # Save to a buffer and return to Lovable
    output = BytesIO()
    prs.save(output)
    return StreamingResponse(BytesIO(output.getvalue()), media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")