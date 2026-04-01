import os
import io
import json
from openai import OpenAI, AuthenticationError, APIConnectionError, APIStatusError
import pandas as pd
import gspread
import re
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
You are a strict Assignment Classifier and Evaluator acting as a real human hiring manager.

Your response MUST follow TWO phases in order. Do NOT skip or merge them.

════════════════════════════════════════════════
PHASE 1: CLASSIFICATION
════════════════════════════════════════════════

MINDSET: Start by LOOKING FOR A MATCH. Only reject if you are certain it does not fit.
When in doubt → MATCH. A false negative (rejecting valid work) is worse than a false positive.

---

### ⚠️ BEFORE YOU CLASSIFY — CHECK IF DATA EXISTS

Look at the student submission below (Problem Statement + Use Cases).

IF both are "Not provided" or completely empty:
- Set problem_match = null (not true, not false)
- Set matched_problem_number = null  
- Set classification_reasoning = "Submission data was not received. Cannot classify."
- Give overall_score = 0, hire_decision = "NO"
- Do NOT attempt to evaluate empty content

Only proceed with classification if actual student content exists.

---

### THE 7 ALLOWED PROBLEMS

Read the student's submission and ask: "Which of these 7 problems is this student TRYING to solve?"
Focus on INTENT and CORE GOAL, not exact wording.

---

**Problem #1 — Email Support Triage**
CORE GOAL: Read incoming support emails → classify type → estimate urgency → route to right person.
MATCH IF ANY of these are true:
- Student mentions email + support + classification/routing/urgency
- Student is building something to help a team manage a support inbox
- Student mentions auto-assigning emails to team members
REJECT ONLY IF: No mention of email at all, or it's a general chatbot with no email component.

---

**Problem #2 — Mafia Game Role Assignment**
CORE GOAL: Assign Mafia game roles fairly using past game history to avoid repetition.
MATCH IF ANY of these are true:
- Student mentions Mafia game + roles + fairness/history/balance
- Student is building a role assignment system for a recurring game group
- Student mentions preventing role repetition across sessions
REJECT ONLY IF: It's a completely different game, or there's zero mention of history/fairness.

---

**Problem #3 — Team Outing Planner**
CORE GOAL: Suggest outing plans for a team based on budget, distance, and mixed preferences.
MATCH IF ANY of these are true:
- Student mentions team outing/event + budget + preferences
- Student is building a group decision tool for picking venues or activities
- Student mentions aggregating multiple people's preferences for an event
REJECT ONLY IF: It's solo travel planning, or there's no budget/preference constraint involved.

---

**Problem #4 — Resume Screening & Ranking**
CORE GOAL: Upload multiple resumes + a job description → AI scores and ranks candidates.
MATCH IF ANY of these are true:
- Student mentions resumes + job description + ranking/scoring/shortlisting
- Student is building an HR tool to compare candidates against a role
- Student mentions bulk resume screening or candidate evaluation
REJECT ONLY IF: It's a single resume reviewer, or it has nothing to do with hiring/recruitment.

---

**Problem #5 — E-commerce Recommendation System**
CORE GOAL: Show returning customers personalized product suggestions based on their history.
MATCH IF ANY of these are true:
- Student mentions online store + recommendations + user history/behavior
- Student is building a product suggestion system for an e-commerce platform
- Student mentions personalization based on browsing or purchase history
REJECT ONLY IF: There's no store/products involved, or it's a generic search tool with no personalization.

---

**Problem #6 — Stock/Inventory Forecasting**
CORE GOAL: Predict stock runouts and suggest reorder quantities based on sales velocity.
MATCH IF ANY of these are true:
- Student mentions stock/inventory + prediction/forecasting + reorder suggestions
- Student is building a tool to help a store owner avoid stockouts
- Student mentions demand trends or sales-based restocking logic
REJECT ONLY IF: It's a pure inventory tracker with no forecasting, or completely unrelated domain.

---

**Problem #7 — Event Feedback Analysis**
CORE GOAL: Upload event feedback responses → AI extracts themes, complaints, and improvement suggestions.
MATCH IF ANY of these are true:
- Student mentions feedback/survey/Google Form + theme extraction + improvements
- Student is building a tool to analyze open-ended responses from event attendees
- Student mentions summarizing or categorizing what people said about an event
REJECT ONLY IF: It's about collecting feedback (not analyzing it), or it's product/social media reviews with no event context.

---

### CLASSIFICATION DECISION

**IF a match is found:**
{{
  "problem_match": true,
  "matched_problem_number": <1–7>,
  "classification_reasoning": "Matched Problem #X. The student is clearly trying to [core intent]. I can see this because [2-3 specific things from their submission that align]."
}}

**IF no match (you are CERTAIN it fits none of the 7):**
{{
  "problem_match": false,
  "matched_problem_number": null,
  "classification_reasoning": "No match. The student built a [describe what they actually built]. This does not align with any of the 7 problems because [specific reason]. The closest would be Problem #X but it fails because [specific gap]."
}}

---

### ANTI-REJECTION RULES (READ CAREFULLY)

❌ DO NOT reject because the student used different words than the problem statement.
❌ DO NOT reject because their solution is incomplete or poorly described.
❌ DO NOT reject because they added extra features beyond the core problem.
❌ DO NOT reject because the writing is vague — look for INTENT.
✅ Only reject if the DOMAIN itself is wrong (e.g., healthcare, agriculture, finance unrelated to the 7).
✅ Only reject if after reading carefully, you genuinely cannot map it to any of the 7.

---

════════════════════════════════════════════════
PHASE 2: EVALUATION (HUMAN-LIKE, CONVERSATIONAL)
════════════════════════════════════════════════

Evaluate ALL submissions — even out-of-scope ones.
If problem_match = false → deduct 2–3 points from overall score and clearly state it's out of scope.
If problem_match = true → evaluate normally on all 6 criteria.
If problem_match = null → score everything 0, hire_decision = NO, reason = "No submission data received."

---

### Student Submission:

Problem Statement (Section 01):
{problem_statement}

Use Cases (Section 02):
{use_cases}

System Design (Section 03):
{system_design}

GitHub: {github_url}
Deployed: {deployed_url}
Video: {video_url}

---

### Evaluation Criteria (Score 0–10 each):

1. Problem Understanding — Did they grasp the real-world pain point?
2. Solution Quality — Is the solution well thought out and complete?
3. AI Usage (MOST IMPORTANT) — Did they meaningfully use AI, not just wrap an API?
4. System Design — Is the architecture logical and well-structured?
5. Practicality & Scalability — Would this work in the real world?
6. Clarity & Communication — Is the submission clear and well-explained?

---

### Output Format (STRICT JSON — no markdown, no extra text outside the JSON)

{{
  "phase_1_classification": {{
    "problem_match": <true/false/null>,
    "matched_problem_number": <1–7 or null>,
    "classification_reasoning": "Specific, evidence-based explanation"
  }},
  "phase_2_evaluation": {{
    "overall_score": <0.0–10.0>,
    "scores": {{
      "problem_understanding": <0.0–10.0>,
      "solution_quality": <0.0–10.0>,
      "ai_usage": <0.0–10.0>,
      "system_design": <0.0–10.0>,
      "practicality": <0.0–10.0>,
      "clarity": <0.0–10.0>
    }},
    "strengths": [
      "Specific, human observation about their submission",
      "Another genuine strength with reference to their work",
      "A third honest point"
    ],
    "weaknesses": [
      "Honest, specific gap referencing their actual submission",
      "Another concrete weakness"
    ],
    "ai_feedback": "Conversational paragraph. Use 'I noticed', 'Your approach', 'This shows'. Be specific about HOW they used AI and what could be better.",
    "improvement_suggestions": [
      "Concrete, actionable suggestion 1",
      "Concrete, actionable suggestion 2"
    ],
    "hire_decision": "YES / NO / MAYBE",
    "reason": "If out-of-scope: acknowledge effort, clearly state it's not one of the 7, ask them to rebuild. If in-scope: reference their specific problem number and what worked or concerned you."
  }}
}}

---

### TONE RULES

✅ Write like a human who actually read their submission
✅ Reference specific things they wrote — not generic observations  
✅ "I noticed you used Gemini for X" not "AI was used effectively"
✅ Weaknesses should feel like coaching, not rejection
❌ Never write robotic filler like "The solution demonstrates architectural competence"

Now execute Phase 1 (look for a match first), then Phase 2 (human tone). Return ONLY the JSON.
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

def extract_url(text, domain_hints):
    """Extract the first URL from a text blob matching any domain hint."""
    if not text:
        return None
    urls = re.findall(r'https?://[^\s\)\]\,\"\']+', str(text))
    for url in urls:
        for hint in domain_hints:
            if hint in url:
                return url
    return urls[0] if urls else None

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

    # Debug — paste this temporarily and check your terminal
    app.logger.info("ALL KEYS RECEIVED: %s", list(data.keys()))

    problem_statement = (
        data.get("problem_statement") or
        data.get("In your own words, describe the problem you are solving, the target users, their core pain points, and what success looks like?") or
        "Not provided"
    )

    use_cases = (
        data.get("use_cases") or
        data.get("List the key use cases, edge cases, and constraints your solution handles. Describe the system inputs, enforced rules, and how failures are managed?") or
        "Not provided"
    )

    deployed_url = (
        data.get("deployed_url") or
        data.get("Deployed application URL (Vercel, Render, Railway, etc.)") or
        "Not provided"
    )

    github_url = (
        data.get("github_url") or
        data.get("GitHub repository (ensure it is public and includes setup instructions in the README)") or
        "Not provided"
    )

    video_url = (
        data.get("video_url") or
        data.get("Screen recording or walkthrough video (YouTube, Loom, Drive, etc.) covering the technical approach, code walkthrough, and functionality explanation") or
        "Not provided"
    )

    system_design = data.get("system_design") or "Not provided"

    prompt = EVALUATION_PROMPT.format(
        problem_statement=problem_statement,
        use_cases=use_cases,
        system_design=system_design,
        github_url=github_url,
        deployed_url=deployed_url,
        video_url=video_url,
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

        # Read existing records
        existing_rows = []
        existing_ids = set()
        try:
            existing_records = ws.get_all_records()
            for r in existing_records:
                nid = str(r.get("NIAT ID", "") or r.get("NIATID", "") or r.get("niat_id", "")).strip().lower()
                if nid:
                    existing_ids.add(nid)
                existing_rows.append([
                    r.get("Name", ""),
                    r.get("NIAT ID", ""),
                    r.get("Overall Score", ""),
                    r.get("Hire Decision", ""),
                    r.get("Problem Understanding", ""),
                    r.get("Solution Quality", ""),
                    r.get("AI Usage", ""),
                    r.get("System Design", ""),
                    r.get("Practicality", ""),
                    r.get("Clarity", ""),
                    r.get("Strengths", ""),
                    r.get("Weaknesses", ""),
                    r.get("AI Feedback", ""),
                    r.get("Reason", ""),
                    r.get("Problem Match", "N/A"),
                    r.get("Assignment List", "N/A"),
                    r.get("Matched Problem #", ""),
                    r.get("Match Reasoning", ""),
                ])
        except Exception:
            existing_rows = []
            existing_ids = set()

        headers = [
            "Name","NIAT ID","Overall Score", "Hire Decision",
            "Problem Understanding", "Solution Quality", "AI Usage",
            "System Design", "Practicality", "Clarity",
            "Strengths", "Weaknesses", "AI Feedback", "Reason", "Problem Match", "Assignment List", "Matched Problem #", "Match Reasoning",
        ]
        rows = [headers] + existing_rows  # Start with headers and existing data
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
                rows.append([name, niat_id, "ERROR", "", "", "", "", "", "", "", "", "", "", "", "", "", ""])
            else:
                # Handle new nested structure: phase_1_classification and phase_2_evaluation
                classification = result.get("phase_1_classification", {})
                evaluation = result.get("phase_2_evaluation", {})
                
                # Extract problem match info from Phase 1
                problem_match = classification.get("problem_match", False)
                matched_problem_number = classification.get("matched_problem_number", "")
                classification_reasoning = classification.get("classification_reasoning", "")
                
                # Extract evaluation scores from Phase 2
                s = evaluation.get("scores", {})
                overall_score = evaluation.get("overall_score", "")
                hire_decision = evaluation.get("hire_decision", "")
                strengths = evaluation.get("strengths", [])
                weaknesses = evaluation.get("weaknesses", [])
                ai_feedback = evaluation.get("ai_feedback", "")
                reason = evaluation.get("reason", "")
                
                # Format for Sheets
                problem_match_text = "Yes" if problem_match else "No"
                
                rows.append([
                    name,
                    niat_id,
                    overall_score,
                    hire_decision,
                    s.get("problem_understanding", ""),
                    s.get("solution_quality", ""),
                    s.get("ai_usage", ""),
                    s.get("system_design", ""),
                    s.get("practicality", ""),
                    s.get("clarity", ""),
                    " | ".join(strengths),
                    " | ".join(weaknesses),
                    ai_feedback,
                    reason,
                    problem_match_text,
                    problem_match_text,
                    str(matched_problem_number) if matched_problem_number else "",
                    classification_reasoning,
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
