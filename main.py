from fastapi import FastAPI, Body
from pydantic import BaseModel
import pandas as pd
from fastapi.responses import JSONResponse
from typing import Optional
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # อนุญาตทุกโดเมน
    allow_credentials=True,
    allow_methods=["*"],  # อนุญาตทุกเมธอด (POST, GET, OPTIONS ฯลฯ)
    allow_headers=["*"],  # อนุญาตทุก Header
)

# Load data from Excel files
path = "/Users/reg/Desktop/admission_weight68_update.xlsx"  # คะแนนที่ใช้คำนวน
path2 = "/Users/reg/Desktop/คณะสาขาที่เปิดรับและคะแนนรวมขั้นต่ำ68-1.xlsx"  # ไฟล์เกณฑ์แต่ละปี
path3 = "/Users/reg/Desktop/คะแนนสูง-ต่ำ2568.xlsx"  # ไฟล์วิเคราะห์คะแนน
path4 = "/Users/reg/Desktop/result2.xlsx"  # ไฟล์วิเคราะห์คะแนน
path5 = "/Users/reg/Desktop/result3.xlsx"  # ไฟล์วิเคราะห์คะแนน
pf = pd.read_excel(path).fillna(0)
pf2 = pd.read_excel(path2).fillna(0)
pf3 = pd.read_excel(path3).fillna(0)
pf4 = pd.read_excel(path4).fillna(0)
pf5 = pd.read_excel(path5).fillna(0)


# Pydantic model to accept the input data
class Scores(BaseModel):
    gpax: Optional[float] = None
    thai_101: Optional[float] = None
    eng_102: Optional[float] = None
    math_103: Optional[float] = None
    sci_201: Optional[float] = None
    chem_202: Optional[float] = None
    bio_203: Optional[float] = None
    phy_204: Optional[float] = None
    fr_011: Optional[float] = None
    de_012: Optional[float] = None
    zh_013: Optional[float] = None
    ja_014: Optional[float] = None
    ko_015: Optional[float] = None
    es_016: Optional[float] = None
    music_021: Optional[float] = None
    exp_visual_art_024: Optional[float] = None
    drawing_023: Optional[float] = None
    commu_drawing_025: Optional[float] = None
    commu_design_026: Optional[float] = None
    ar_041: Optional[float] = None
    design_042: Optional[float] = None
    art_051: Optional[float] = None
    physical_052: Optional[float] = None
    tech_for_med_vision_061: Optional[float] = None
    art_for_med_vision_062: Optional[float] = None
    tpat3_30: Optional[float] = None
    tgat_90: Optional[float] = None
    tgat1_91: Optional[float] = None
    tgat2_92: Optional[float] = None
    tgat3_93: Optional[float] = None


# Helper function to convert score from "N" to None
def convert_score(value):
    if value is None or value == "N":
        return None
    return float(value)


@app.post("/programs", response_class=JSONResponse)
async def get_qualified_programs(data: Scores):
    
    # Rest of the code remains the same
    score = {
        "101": data.thai_101,
        "102": data.eng_102,
        "103": data.math_103,
        "201": data.sci_201,
        "202": data.chem_202,
        "203": data.bio_203,
        "204": data.phy_204,
        "11": data.fr_011,
        "12": data.de_012,
        "13": data.zh_013,
        "14": data.ja_014,
        "15": data.ko_015,
        "16": data.es_016,
        "21": data.music_021,
        "24": data.exp_visual_art_024,
        "23": data.drawing_023,
        "25": data.commu_drawing_025,
        "26": data.commu_design_026,
        "41": data.ar_041,
        "42": data.design_042,
        "51": data.art_051,
        "52": data.physical_052,
        "61": data.tech_for_med_vision_061,
        "62": data.art_for_med_vision_062,
        "30": data.tpat3_30,
        "90": data.tgat_90,
        "91": data.tgat1_91,
        "92": data.tgat2_92,
        "93": data.tgat3_93,
    }

    program_status = {}
    program_requirements = {}
    
    # Extract program requirements from pf2
    for _, row in pf2.iterrows():
        program_id = row["WPROGRAMID"]
        program_requirements[program_id] = {
            "min_score": row["MIN_SCORE"],
            "gpax_require": row["GPAX_REQUIRE"],
        }

    # Calculate the program status based on the provided scores
    for _, row in pf.iterrows():
        faculty_name = row["FACULTYNAME"]
        program_id = row["WPROGRAMID"]
        program_name = row["PROGRAMNAME"]
        subject_code = str(row["SUBJECTCODE2"])
        program_key = (faculty_name, program_id, program_name)

        if program_key not in program_status:
            program_status[program_key] = {"sum_weight": 0, "pass_all": True}

        if subject_code in score and score[subject_code] is not None:
            if score[subject_code] >= row["MINSCORE"]:
                program_status[program_key]["sum_weight"] += score[subject_code] * (
                    row["WEIGHT"] / 100
                )
            else:
                program_status[program_key]["pass_all"] = False
        else:
            program_status[program_key]["pass_all"] = False

    # Check which programs qualify
    qualified_programs = []
    for (faculty_name, program_id, program_name), status in program_status.items():
        if program_id in program_requirements:
            min_score_required = program_requirements[program_id]["min_score"]
            gpax_required = program_requirements[program_id]["gpax_require"]

            if (
                status["pass_all"]
                and status["sum_weight"] >= min_score_required
                and data.gpax >= gpax_required
            ):
                qualified_programs.append(
                    {
                        "faculty": faculty_name,
                        "program_id": program_id,
                        "program_name": program_name,
                        "total_score": round(status["sum_weight"], 3),
                        "min_score": min_score_required,
                        "gpax_required": gpax_required
                    }
                )

    # Sort the programs by total score (descending)
    qualified_programs = sorted(
        qualified_programs, key=lambda x: x["total_score"], reverse=True
    )

    return qualified_programs


@app.post("/list_programs", response_class=JSONResponse)
async def list_name_programs():
    list_programs = []
    id = 0

    seen_programs = set()
    for _, row in pf3.iterrows():
        program_id = row["PROGRAMID"]
        program_name = row["PROGRAMNAME"]
        min_score = row["MIN_SCORE"]
        max_score = row["MAX_SCORE"]

        if (program_id, program_name) not in seen_programs:
            id += 1
            faculty_name = row["FACULTYNAME"]
            seen_programs.add((program_id, program_name))

            list_programs.append(
                {
                    "id": id,
                    "faculty_name": faculty_name,
                    "program_id": program_id,
                    "program_name": program_name,
                    "min_score": min_score,
                    "max_score": max_score,
                }
            )

    return list_programs


@app.post("/list_programs_zscore", response_class=JSONResponse)
async def list_name_programs():
    list_programs = []
    id = 0

    seen_programs = set()
    for _, row in pf5.iterrows():
        program_id = row["PROGRAMID"]
        program_name = row["PROGRAMNAME"]
        mean_avg = row["mean_avg"]
        sd_avg = row["sd_avg"]

        if (program_id, program_name) not in seen_programs:
            id += 1
            faculty_name = row["FACULTYNAME"]
            seen_programs.add((program_id, program_name))

            list_programs.append(
                {
                    "id": id,
                    "faculty_name": faculty_name,
                    "program_id": program_id,
                    "program_name": program_name,
                    "mean_avg": mean_avg,
                    "sd_avg": sd_avg,
                }
            )

    return list_programs


@app.post("/select_list_programs", response_class=JSONResponse)
async def select_list_programs():
    list_programs = []
    id = 0

    # ใช้ dictionary เก็บข้อมูล program และ subjects
    programs_dict = {}

    for _, row in pf.iterrows():
        program_id = row["WPROGRAMID"]
        program_name = row["PROGRAMNAME"]
        faculty = row["FACULTYNAME"]
        subject_code = str(row["SUBJECTCODE2"])
        subject_name = row["SUBJECTNAME"]

        # สร้าง key สำหรับแต่ละ program
        program_key = (program_id, program_name, faculty)

        # ถ้า program ยังไม่เคยเจอ ให้สร้างใหม่
        if program_key not in programs_dict:
            programs_dict[program_key] = {
                "program_id": program_id,
                "program_name": program_name,
                "faculty": faculty,
                "subject_codes": [],
                "subject_names": [],
            }

        # เพิ่ม subject เข้าไปใน list (ตรวจสอบไม่ให้ซ้ำ)
        if subject_code not in programs_dict[program_key]["subject_codes"]:
            programs_dict[program_key]["subject_codes"].append(subject_code)
            programs_dict[program_key]["subject_names"].append(subject_name)

    # แปลง dictionary กลับเป็น list
    for program_info in programs_dict.values():
        id += 1
        list_programs.append(
            {
                "id": id,
                "faculty": program_info["faculty"],
                "program_id": program_info["program_id"],
                "program_name": program_info["program_name"],
                "subject_codes": program_info["subject_codes"],  # เป็น array
                "subject_names": program_info["subject_names"],  # เป็น array
            }
        )

    return list_programs

@app.post("/programs_list", response_class=JSONResponse)
async def combined_programs_list():
    list_programs = []
    id = 0

    # Step 1: ดึงข้อมูล z-score จาก pf5
    zscore_mapping = {}
    seen_programs = set()
    
    for _, row in pf5.iterrows():
        program_id = row["PROGRAMID"]
        program_name = row["PROGRAMNAME"]
        program_name_trimmed = str(program_name).replace(" ", "").strip()  # ลบช่องว่างทั้งหมด
        mean_avg = row["mean_avg"]
        sd_avg = row["sd_avg"]
        faculty_name = row["FACULTYNAME"]

        if (program_id, program_name) not in seen_programs:
            seen_programs.add((program_id, program_name))
            zscore_mapping[program_name_trimmed] = {
                "program_id": program_id,
                "faculty_name": faculty_name,
                "mean_avg": mean_avg,
                "sd_avg": sd_avg,
                "original_name": program_name
            }

    # Step 2: ดึงข้อมูล subjects จาก pf และ match กับ z-score
    programs_dict = {}

    for _, row in pf.iterrows():
        program_id = row["WPROGRAMID"]
        program_name = row["PROGRAMNAME"]
        program_name_trimmed = str(program_name).replace(" ", "").strip()
        faculty = row["FACULTYNAME"]
        subject_code = str(row["SUBJECTCODE2"])
        subject_name = row["SUBJECTNAME"]
        weight = row["WEIGHT"]

        program_key = (program_id, program_name_trimmed, faculty)

        if program_key not in programs_dict:
            programs_dict[program_key] = {
                "program_id": program_id,
                "program_name": program_name,
                "program_name_trimmed": program_name_trimmed,
                "faculty": faculty,
                "subject_codes": [],
                "subject_names": [],
                "weights": [],
                "max_weight_subjects": []  # New field for subjects with max weight
            }

        if subject_code not in programs_dict[program_key]["subject_codes"]:
            programs_dict[program_key]["subject_codes"].append(subject_code)
            programs_dict[program_key]["subject_names"].append(subject_name)
            programs_dict[program_key]["weights"].append(weight)

            # Update max weight subjects
            current_max = max(programs_dict[program_key]["weights"])
            if weight == current_max:
                programs_dict[program_key]["max_weight_subjects"].append(subject_name)
            elif weight > current_max:
                programs_dict[program_key]["max_weight_subjects"] = [subject_name]

    # Step 3: รวมข้อมูลและ match โดยใช้ program_name ที่ trim แล้ว
    for program_info in programs_dict.values():
        id += 1
        
        program_name_trimmed = program_info["program_name_trimmed"]
        zscore_data = zscore_mapping.get(program_name_trimmed, {
            "program_id": None,
            "faculty_name": program_info["faculty"],
            "mean_avg": 0,
            "sd_avg": 0,
            "original_name": program_info["program_name"]
        })

        list_programs.append({
            "id": id,
            "faculty": program_info["faculty"],
            "program_id": program_info["program_id"],
            "program_name": program_info["program_name"],
            "program_name_trimmed": program_name_trimmed,
            "subject_codes": program_info["subject_codes"],
            "subject_names": program_info["subject_names"],
            "weights": program_info["weights"],
            "max_weight_subjects": program_info["max_weight_subjects"],  # Added max weight subjects
            "mean_avg": zscore_data["mean_avg"],
            "sd_avg": zscore_data["sd_avg"],
            "total_subjects": len(program_info["subject_codes"]),
            "has_zscore_data": zscore_data["mean_avg"] > 0,
            "matched_with": zscore_data.get("original_name", "No match")
        })

    list_programs = sorted(list_programs, key=lambda x: (x["faculty"], x["program_name"]))
    
    for i, program in enumerate(list_programs, 1):
        program["id"] = i

    return list_programs
