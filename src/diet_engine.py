import json
import copy
import os

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")

def load_food_db():
    path = os.path.join(DATA_DIR, "food_database.json")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def load_default_plan():
    path = os.path.join(DATA_DIR, "default_weekly_plan.json")
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def apply_adjustments(weekly_plan, options):
    """
    options = {
        'complaint': 'constipation' | 'plateau' | 'none',
        'breakfast_protein': 'eggs' | 'cottage_cheese' | 'foul' | 'default',
        'open_meal_day': 'الجمعة' | 'الخميس' | 'none',
        'detox_day': 'السبت' | 'none'
    }
    """
    plan = copy.deepcopy(weekly_plan)
    complaint = options.get("complaint", "none")
    b_protein = options.get("breakfast_protein", "default")
    open_meal_day = options.get("open_meal_day", "none")
    detox_day = options.get("detox_day", "none")

    # 1. Modify Breakfast protein if requested
    if b_protein == "cottage_cheese":
        for day in plan:
            for meal in day["meals"]:
                if "إفطار" in meal["meal_type"] or "الإفطار" in meal["meal_type"]:
                    meal["content"] = meal["content"].replace("1 بيضة (مسلوقة أو مقلية)", "4 ملاعق جبن قريش بزيت الزيتون")
                    meal["content"] = meal["content"].replace("2 بيضة مسلوقة أو مقلية", "4 ملاعق جبن قريش مع زعتر")
                    meal["content"] = meal["content"].replace("Omelet  (2 بيضه + إضافات الخضار والمشروم حسب الرغبة)", "جبن قريش مع قطع الخيار والطماطم وزيت الزيتون")
    elif b_protein == "foul":
        for day in plan:
            for meal in day["meals"]:
                if "إفطار" in meal["meal_type"] or "الإفطار" in meal["meal_type"]:
                    meal["content"] = meal["content"].replace("1 بيضة (مسلوقة أو مقلية)", "4 إلى 5 ملاعق فول بزيت الزيتون والليمون")

    # 2. Open Meal Day
    if open_meal_day and open_meal_day != "none":
        for day in plan:
            if day["day"] == open_meal_day:
                for meal in day["meals"]:
                    if "غداء" in meal["meal_type"] or "عشاء" in meal["meal_type"] or "العشاء" in meal["meal_type"]:
                        meal["content"] = "🍔 وجبة مفتوحة (Open Meal) حسب الرغبة باعتدال + 2 كوب ماء كبير قبلها"
                        break

    # 3. Detox Day
    if detox_day and detox_day != "none":
        for day in plan:
            if day["day"] == detox_day:
                for meal in day["meals"]:
                    if "خفيفة" in meal["meal_type"]:
                        meal["content"] = "مشروب الديتوكس الأخضر (خيار + كرفس + تفاح أخضر + ليمون ونعناع) أو شاي أخضر بالزنجبيل"
                    elif "عشاء" in meal["meal_type"] or "العشاء" in meal["meal_type"]:
                        meal["content"] = "طبق ديتوكس سلطة خضراء غنية بالورقيات + 1 كوب زبادي لايت مع بذور الشيا"

    # 4. Complaints: Constipation (إمساك)
    if complaint == "constipation":
        for day in plan:
            for meal in day["meals"]:
                if "خفيفة" in meal["meal_type"]:
                    if "بذور الشيا" not in meal["content"]:
                        meal["content"] += " + إضافة 1 ملعقة صغيرة بذور شيا مع كوب ماء دافئ أو قراصيا (برقوق مجفف)"

    # 5. Complaints: Weight Plateau (ثبات وزن)
    if complaint == "plateau":
        for idx, day in enumerate(plan):
            if idx in [1, 3]:  # Sunday & Tuesday: Low Carb
                for meal in day["meals"]:
                    if "عشاء" in meal["meal_type"] or "العشاء" in meal["meal_type"]:
                        meal["content"] += " (وجبة خالية تماماً من النشويات لكسر ثبات الوزن)"

    return plan
