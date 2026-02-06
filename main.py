from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import StreamingResponse
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
                if not shape.has_text_frame: 
                    continue
                
                tf = shape.text_frame
                # Target only the first paragraph and first run to preserve 'SLM' design
                if tf.paragraphs:
                    p = tf.paragraphs[0]
                    
                    if shape == slide.shapes.title:
                        new_text = slide_update.get('new_title', '')
                    else:
                        new_text = "\n".join(slide_update.get('new_content', []))

                    # SURGICAL SWAP: Replace text in the first run to keep formatting
                    if p.runs:
                        p.runs[0].text = new_text
                        # Remove any extra runs that might contain leftover original text
                        for extra_run in p.runs[1:]:
                            p._p.remove(extra_run._r) 
                    else:
                        # Fallback if the template shape was empty
                        p.text = new_text

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    
    return StreamingResponse(
        output, 
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=Rebuilt_Presentation.pptx"}
    )