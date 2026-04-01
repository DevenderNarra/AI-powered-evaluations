import os
import io
import json
from openai import OpenAI, AuthenticationError, APIConnectionError, APIStatusError
import pandas as pd
import gspread
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, send_file
from dotenv import load_dotenv
load_dotenv()

SPREADSHEET_ID = "1CFviUlaAhBqV_cE1SgnZsgOriLYbx4_s1G6MOjfrdaI"
INPUT_SHEET_GID = 1267953312
OUTPUT_SHEET_GID = 527486028


def get_gspread_client():
    creds_file = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
    token_file = os.environ.get("GOOGLE_TOKEN_FILE", "token.json")
    # gspread.oauth handles browser login on first run and token refresh automatically
    return gspread.oauth(
        credentials_filename=creds_file,
        authorized_user_filename=token_file,
    )


app = Flask(__name__)

EVALUATION_PROMPT = """
You are a Senior AI Product Engineer and Hiring Evaluator.

Your task is to evaluate a student's project submission based on real-world product thinking, AI usage, and engineering quality.

---

### 📌 Evaluation Context

The student was given the following instructions:

- Pick one problem statement from the given list
- Understand the problem deeply before coding
- Identify users and pain points
- Define success criteria
- Build an AI-powered solution (NLP / ML / Generative AI)
- Ensure the solution is practical, simple, and scalable

---

### 📋 Allowed Problem Statements

Students must choose from these 7 problems:

#1: AI system that reads incoming support emails, identifies issue type, estimates urgency, and suggests team member assignment.

#2: AI system for Mafia game role assignment that checks history and ensures variety/balance.

#3: AI-powered outing planner that suggests plans based on team preferences, budget, distance, and food preferences.

#4: AI system for resume screening that scores and ranks candidates based on job description match.

#5: AI-powered recommendation system for online store that shows relevant products to returning customers.

#6: AI-powered stock management system that predicts stock depletion and suggests reorder amounts.

#7: AI-powered feedback analysis tool that identifies themes, highlights complaints, and suggests improvements.

---

### 📥 Student Submission:

Problem Statement & Users:
{problem_statement}

Use Cases & Edge Cases:
{use_cases}

System Design (Inputs, Rules, Failures):
{system_design}

GitHub:
{github_url}

Deployment:
{deployed_url}

Video:
{video_url}

---

### 📊 Evaluation Criteria (Score each out of 10, use decimal values like 6.5, 7.0, 8.5):

1. Problem Understanding
2. Solution Quality
3. AI Usage (VERY IMPORTANT)
4. System Design Thinking
5. Practicality & Scalability
6. Clarity & Communication

---

### 🧾 Output Format (STRICT JSON ONLY)

IMPORTANT: All scores must be numbers between 0.0 and 10.0. Do NOT use 0-100 scale.

{{
  "overall_score": <number between 0.0 and 10.0>,
  "scores": {{
    "problem_understanding": <number between 0.0 and 10.0>,
    "solution_quality": <number between 0.0 and 10.0>,
    "ai_usage": <number between 0.0 and 10.0>,
    "system_design": <number between 0.0 and 10.0>,
    "practicality": <number between 0.0 and 10.0>,
    "clarity": <number between 0.0 and 10.0>
  }},
  "strengths": ["...", "..."],
  "weaknesses": ["...", "..."],
  "ai_feedback": "...",
  "improvement_suggestions": ["...", "..."],
  "hire_decision": "YES / NO / MAYBE",
  "reason": "..."
}}

---

### ⚠️ Rules:
- Penalize submissions that do not use one of the allowed problem statements
- Penalize weak or fake AI usage
- Penalize missing GitHub / Deployment
- Be strict like a hiring manager
- Avoid generic feedback
- Return ONLY the JSON object, no markdown, no explanation

---

Now evaluate this submission.
"""


EXPECTED_COLUMNS = [
    "Student Name",
    "Problem Statement",
    "Use Cases",
    "System Design",
    "GitHub URL",
    "Deployed URL",
    "Video URL",
]


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/parse-excel", methods=["POST"])
def parse_excel():
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file uploaded."}), 400

    try:
        df = pd.read_excel(file, engine="openpyxl")
    except Exception as e:
        return jsonify({"error": f"Could not read file: {str(e)}"}), 400

    df = df.fillna("")
    rows = df.to_dict(orient="records")
    columns = list(df.columns)
    return jsonify({"columns": columns, "rows": rows})


@app.route("/download-template")
def download_template():
    df = pd.DataFrame(columns=EXPECTED_COLUMNS)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Submissions")
    output.seek(0)
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="submission_template.xlsx",
    )


@app.route("/evaluate", methods=["POST"])
def evaluate():
    data = request.json
    prompt = EVALUATION_PROMPT.format(
        problem_statement=data.get("problem_statement", "Not provided"),
        use_cases=data.get("use_cases", "Not provided"),
        system_design=data.get("system_design", "Not provided"),
        github_url=data.get("github_url", "Not provided"),
        deployed_url=data.get("deployed_url", "Not provided"),
        video_url=data.get("video_url", "Not provided"),
    )

    def generate():
        client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        full_text = ""
        error_msg = None

        try:
            stream = client.chat.completions.create(
                model="gpt-4o-mini",
                max_tokens=4096,
                messages=[{"role": "user", "content": prompt}],
                stream=True,
            )
            for chunk in stream:
                text = chunk.choices[0].delta.content or ""
                if text:
                    full_text += text
                    yield f"data: {json.dumps({'chunk': text})}\n\n"
        except AuthenticationError:
            error_msg = "Invalid API key. Check OPENAI_API_KEY in your .env file."
        except APIConnectionError:
            error_msg = "Network error. Check your internet connection."
        except APIStatusError as e:
            error_msg = f"API error {e.status_code}: {e.message}"
        except Exception as e:
            error_msg = f"Unexpected error: {str(e)}"

        if error_msg:
            yield f"data: {json.dumps({'error': error_msg})}\n\n"
        else:
            yield f"data: {json.dumps({'done': True, 'full': full_text})}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/load-sheets")
def load_sheets():
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(SPREADSHEET_ID)
        ws = sh.get_worksheet_by_id(INPUT_SHEET_GID)

        # Load already evaluated NIAT IDs from output sheet to avoid duplicates
        existing_ids = set()
        try:
            out_ws = sh.get_worksheet_by_id(OUTPUT_SHEET_GID)
            out_rows = out_ws.get_all_records()
            for r in out_rows:
                nid = str(r.get("NIAT ID", "") or r.get("NIATID", "") or r.get("niat_id", "")).strip().lower()
                if nid:
                    existing_ids.add(nid)
        except Exception:
            existing_ids = set()

        records = ws.get_all_records()
        skipped = 0
        filtered = []
        if records:
            for row in records:
                nid = str(row.get("NIAT ID", "") or row.get("NIATID", "") or row.get("niat_id", "")).strip().lower()
                if nid and nid in existing_ids:
                    skipped += 1
                    continue
                filtered.append(row)
            columns = list(records[0].keys())
            return jsonify({"columns": columns, "rows": filtered, "skipped": skipped})

        # Fallback when get_all_records() returns empty (e.g., merged titles/blank leading rows)
        values = ws.get_all_values()
        if not values or len(values) <= 1:
            return jsonify({"columns": [], "rows": [], "skipped": 0})

        header = values[0]
        rows = []
        for row_values in values[1:]:
            if not any(str(v).strip() for v in row_values):
                continue
            row = {header[i] if i < len(header) else f"Column_{i+1}": row_values[i] if i < len(row_values) else "" for i in range(len(header))}
            nid = str(row.get("NIAT ID", "") or row.get("NIATID", "") or row.get("niat_id", "")).strip().lower()
            if nid and nid in existing_ids:
                skipped += 1
                continue
            rows.append(row)

        if not rows:
            return jsonify({"columns": [], "rows": [], "skipped": skipped})

        return jsonify({"columns": header, "rows": rows, "skipped": skipped})

    except Exception as e:
        app.logger.exception("Error loading sheets")
        return jsonify({"error": str(e)}), 500


@app.route("/save-to-sheets", methods=["POST"])
def save_to_sheets():
    data = request.json
    try:
        gc = get_gspread_client()
        sh = gc.open_by_key(SPREADSHEET_ID)
        ws = sh.get_worksheet_by_id(OUTPUT_SHEET_GID)

        existing_ids = set()
        try:
            out_rows = ws.get_all_records()
            for r in out_rows:
                nid = str(r.get("NIAT ID", "") or r.get("NIATID", "") or r.get("niat_id", "")).strip().lower()
                if nid:
                    existing_ids.add(nid)
        except Exception:
            existing_ids = set()

        headers = [
            "Name","NIAT ID","Overall Score", "Hire Decision",
            "Problem Understanding", "Solution Quality", "AI Usage",
            "System Design", "Practicality", "Clarity",
            "Strengths", "Weaknesses", "AI Feedback", "Reason",
        ]
        rows = [headers]
        added = 0

        for item in data:
            student = item.get("student", {})
            result = item.get("result")
            name = student.get("Student Name") or student.get("Name") or ""
            niat_id = student.get("NIAT ID") or student.get("NIATID") or student.get("niat_id") or ""
            nid_norm = str(niat_id).strip().lower()

            if nid_norm and nid_norm in existing_ids:
                continue

            if not result:
                rows.append([name, niat_id, "ERROR", "", "", "", "", "", "", "", "", "", "", ""])
            else:
                s = result.get("scores", {})
                rows.append([
                    name,
                    niat_id,
                    result.get("overall_score", ""),
                    result.get("hire_decision", ""),
                    s.get("problem_understanding", ""),
                    s.get("solution_quality", ""),
                    s.get("ai_usage", ""),
                    s.get("system_design", ""),
                    s.get("practicality", ""),
                    s.get("clarity", ""),
                    " | ".join(result.get("strengths", [])),
                    " | ".join(result.get("weaknesses", [])),
                    result.get("ai_feedback", ""),
                    result.get("reason", ""),
                ])

            if nid_norm:
                existing_ids.add(nid_norm)
            added += 1

        ws.clear()
        ws.update("A1", rows)
        return jsonify({"success": True, "count": added})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)
