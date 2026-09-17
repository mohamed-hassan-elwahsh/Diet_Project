from fastapi import FastAPI, HTTPException, Form
from fastapi.responses import HTMLResponse, FileResponse
from pydantic import BaseModel
import os
import datetime
import json
import re
import subprocess
import google.generativeai as genai
import docx
from diet_engine import load_default_plan, apply_adjustments
from docx_generator import generate_patient_diet_docx

app = FastAPI(title="Diet Assistant AI")

OUTPUT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
os.makedirs(OUTPUT_DIR, exist_ok=True)
DATA_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "data"))
MASTER_PLAN_TXT = os.path.join(DATA_DIR, "master_meal_plan.txt")

PATIENTS_DIR = r"C:\Users\Gobran_Group\OneDrive\Patients"
BACKUP_DIR = r"C:\Users\Gobran_Group\OneDrive\Back-up Diet Plan"

class GenerateRequest(BaseModel):
    patient_name: str
    plan_date: str = ""
    weekly_plan: list = None

def extract_all_text(filepath):
    doc = docx.Document(filepath)
    texts = [node.text.strip() for node in doc.element.xpath('//w:t') if node.text.strip()]
    return " | ".join(texts)

def extract_plan_from_docx(filepath):
    try:
        doc = docx.Document(filepath)
        days_data = []
        day_names = ['السبت', 'الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة']
        
        # Filter to only include actual diet tables (skip intro/header tables)
        valid_tables = []
        for t in doc.tables:
            # A diet table usually has at least 2 rows (header + meal) and 4 columns
            if len(t.rows) > 1 and len(t.rows[0].cells) >= 4:
                valid_tables.append(t)
                
        for t_idx, t in enumerate(valid_tables[:7]):
            day_info = {
                'day': day_names[t_idx] if t_idx < len(day_names) else f'Day {t_idx+1}',
                'meals': []
            }
            for r in t.rows[1:]:
                cells = [c.text.strip() for c in r.cells]
                meal_type = cells[1] if len(cells) > 1 else ''
                timing = cells[2] if len(cells) > 2 else ''
                diet_content = cells[3] if len(cells) > 3 else ''
                # Only add if it looks like a real meal row
                if meal_type or diet_content:
                    day_info['meals'].append({
                        'meal_type': meal_type,
                        'timing': timing,
                        'content': diet_content
                    })
            if day_info['meals']:
                days_data.append(day_info)
        
        if len(days_data) == 7:
            return days_data
    except Exception as e:
        print("Error parsing old diet docx tables:", e)
    
    return load_default_plan()

def get_master_plan_text():
    if os.path.exists(MASTER_PLAN_TXT):
        with open(MASTER_PLAN_TXT, 'r', encoding='utf-8') as f:
            return f.read()
    return ""

@app.get("/api/pick-file")
def pick_file():
    script = f"""
import tkinter as tk
from tkinter import filedialog
import sys
root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)
res = filedialog.askopenfilename(initialdir=r'{PATIENTS_DIR}', filetypes=[("Word Documents", "*.docx")])
print(res)
"""
    try:
        res = subprocess.check_output(['python', '-c', script], text=True, stderr=subprocess.DEVNULL).strip()
        return {"path": res}
    except:
        return {"path": ""}

@app.get("/api/pick-folder")
def pick_folder():
    script = f"""
import tkinter as tk
from tkinter import filedialog
import sys
root = tk.Tk()
root.withdraw()
root.attributes('-topmost', True)
res = filedialog.askdirectory(initialdir=r'{BACKUP_DIR}')
print(res)
"""
    try:
        res = subprocess.check_output(['python', '-c', script], text=True, stderr=subprocess.DEVNULL).strip()
        return {"path": res}
    except:
        return {"path": ""}

@app.post("/api/analyze-patient")
async def analyze_patient(
    api_key: str = Form(None), 
    info_path: str = Form(None),
    diet_folder_path: str = Form(None)
):
    try:
        if not api_key:
            return {"success": False, "error": "يرجى إدخال مفتاح Gemini API Key في أعلى الصفحة."}
        if not info_path or not os.path.exists(info_path):
            return {"success": False, "error": "مسار ملف بيانات المريض غير صحيح أو الملف غير موجود."}

        genai.configure(api_key=api_key)
        
        old_diet_path = None
        if diet_folder_path and os.path.isdir(diet_folder_path):
            latest_file = None
            max_time = 0
            for f in os.listdir(diet_folder_path):
                if f.endswith('.docx') and not f.startswith('~$'):
                    full_p = os.path.join(diet_folder_path, f)
                    mtime = os.path.getmtime(full_p)
                    if mtime > max_time:
                        max_time = mtime
                        latest_file = full_p
            old_diet_path = latest_file
            
        raw_text = extract_all_text(info_path)
        
        if old_diet_path:
            base_plan = extract_plan_from_docx(old_diet_path)
        else:
            base_plan = load_default_plan()
            
        formatted_base_plan = []
        for day in base_plan:
            formatted_day = {"day": day["day"], "meals": []}
            for m in day["meals"]:
                formatted_day["meals"].append({
                    "meal_type": m["meal_type"],
                    "timing": m["timing"],
                    "original_content": m.get("content", ""),
                    "suggested_content": "",
                    "hint": ""
                })
            formatted_base_plan.append(formatted_day)
            
        master_plan_db = get_master_plan_text()
        
        model = genai.GenerativeModel('gemini-3.6-flash')
        prompt = f"""
أنت مساعد ذكي ومحترف لدكتور تغذية.
إليك النص الخام المستخرج من ملف المريض (يحتوي على بياناته وتاريخ جلساته والـ Complaint أو تعليمات الدكتور):
"{raw_text}"

وإليك خطة الوجبات (JSON) للدايت القديم للمريض:
{json.dumps(formatted_base_plan, ensure_ascii=False)}

========================
[قاعدة بيانات الماستر بلان Master Meal Plan]
هذه هي الوجبات المتوفرة لديك للاختيار منها كبدائل مقسمة لجداول وصفوف:
{master_plan_db}
========================

التعليمات الهامة بدقة:
1. استخرج "اسم المريض".
2. ابحث في نص المريض واستخرج بيانات **آخر 5 جلسات** زمنياً (تاريخ كل جلسة والشكوى/الملاحظات الخاصة بها).
3. بالنسبة للوجبات: وفر **من 3 إلى 5 اقتراحات (بدائل)** لكل وجبة رئيسية.
4. **كيفية عمل الاقتراحات (هام جداً):** الاقتراحات يجب أن تكون **شبه الوجبة الأصلية** في الهيكل. يجب عليك الحفاظ على "الثوابت" الموجودة في الدايت القديم (مثل: عبارة "2 كوب ماء قبل الأكل"، "سلطات"، "خضار سوتيه" أو أي إضافات ثابتة). قم **فقط بتغيير المكون الرئيسي** (مثل تبديل الدجاج باللحم أو السمك أو الكوردون بلو) واستخرج هذا المكون البديل من [قاعدة بيانات الماستر بلان]. 
5. يجب أن تكون البدائل مكافئة في القيمة الغذائية وتتوافق مع الشكوى. ولكل بديل املأ `content` (النص النهائي بعد دمج المكون الجديد مع الثوابت القديمة) و `source` (مكان المكون في الماستر بلان).
6. **احذّر من نسيان أي يوم:** يجب أن يكون الرد يحتوي على **الـ 7 أيام كاملة بالترتيب (تبدأ من السبت وتنتهي بالجمعة)** دون حذف أي يوم أو أي وجبة.

أخرج الرد بصيغة JSON فقط، مطابق لهذا الهيكل بالضبط:
{{
  "patient_name": "اسم المريض المستخرج",
  "recent_sessions": [
    {{ "date": "تاريخ الجلسة", "notes": "الشكوى والتفاصيل" }}
  ],
  "suggested_plan": [
    {{
      "day": "السبت",
      "meals": [
        {{
          "meal_type": "...",
          "timing": "...",
          "original_content": "...",
          "suggestions": [
             {{ "content": "محتوى الاقتراح الأول", "source": "جدول رقم X، صف Y" }},
             {{ "content": "محتوى الاقتراح الثاني (بروتين مختلف للروتين)", "source": "جدول رقم X، صف Z" }}
          ]
        }}
      ]
    }}
  ]
}}
لا تكتب أي نص قبل أو بعد الـ JSON.
"""
        response = model.generate_content(prompt)
        txt = response.text.strip()
        
        if txt.startswith("```json"): txt = txt[7:]
        elif txt.startswith("```"): txt = txt[3:]
        if txt.endswith("```"): txt = txt[:-3]
        
        ai_data = json.loads(txt.strip())
        
        return {
            "success": True,
            "patient_name": ai_data.get("patient_name", ""),
            "recent_sessions": ai_data.get("recent_sessions", []),
            "suggested_plan": ai_data.get("suggested_plan", [])
        }
    except Exception as e:
        import traceback
        traceback.print_exc()
        return {"success": False, "error": str(e)}

@app.get("/api/default-plan")
def get_default_plan():
    return []

@app.post("/api/generate-doc")
def generate_doc(req: GenerateRequest):
    if not req.patient_name.strip():
        raise HTTPException(status_code=400, detail="يرجى إدخال اسم المريض")
    
    plan_date = req.plan_date.strip() or datetime.date.today().strftime("%d-%m-%Y")
    
    plan_to_use = []
    if req.weekly_plan:
        for day in req.weekly_plan:
            new_day = {"day": day.get("day", ""), "meals": []}
            for m in day.get("meals", []):
                new_day["meals"].append({
                    "meal_type": m.get("meal_type", ""),
                    "timing": m.get("timing", ""),
                    "content": m.get("selected_content", m.get("original_content", "")),
                    "source": m.get("selected_source", "")
                })
            plan_to_use.append(new_day)
    
    if not plan_to_use:
        plan_to_use = load_default_plan()
    
    import time
    unique_id = str(int(time.time()))[-5:]
    safe_name = "".join(c for c in req.patient_name if c.isalnum() or c in (' ', '-', '_')).strip()
    filename = f"{safe_name}_{plan_date}_{unique_id}.docx"
    filepath = os.path.join(OUTPUT_DIR, filename)
    
    try:
        generate_patient_diet_docx(req.patient_name, plan_date, plan_to_use, filepath)
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"فشل في استخراج الملف: {str(e)}")
        
    return {
        "success": True,
        "filename": filename,
        "download_url": f"/download/{filename}",
        "full_path": filepath
    }

@app.get("/download/{filename}")
def download_file(filename: str):
    path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(path):
        return FileResponse(
            path, 
            filename=filename, 
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Cache-Control": "no-store, no-cache, must-revalidate, max-age=0",
                "Pragma": "no-cache",
                "Expires": "0"
            }
        )
    raise HTTPException(status_code=404, detail="File not found")

@app.get("/", response_class=HTMLResponse)
def index():
    html_path = os.path.join(os.path.dirname(__file__), "index.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return f.read()
    return "<h1>Diet Assistant App</h1>"

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
