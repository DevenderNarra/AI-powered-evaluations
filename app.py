import os
import io
import json
from openai import OpenAI, AuthenticationError, APIConnectionError, APIStatusError
import pandas as pd
import gspread
import re
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, send_file
from dotenv import load_dotenv
import requests
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
You are a STRICT Assignment Classifier and Evaluator acting as a senior technical hiring manager.

Your response MUST follow TWO phases in order. Do NOT skip or merge them.

════════════════════════════════════════════════
MANDATORY 5-SECTION SUBMISSION STRUCTURE
════════════════════════════════════════════════

Every assignment MUST include ALL 5 sections:
1. Your Understanding of the Assignment (Problem Statement) — Demonstrate problem understanding
2. Cases & Logic Constraints (Use Cases) — Map important cases and constraints
3. Technical Approach & Architecture (System Design) — Describe the design and AI usage
4. Code Submission (GitHub) — Working code with complete README
5. Working Demonstration — Deployed URL OR video showing the AI/automation in action

If ANY of the 5 sections are missing or insufficient:
- GitHub missing or private → DEDUCT 2 points from overall score + NOTE IT
- Deployed URL missing → DEDUCT 2.5 points from overall score + NOTE IT
- Video URL missing (and no deployed URL) → DEDUCT 2.5 points from overall score + NOTE IT
- README incomplete/missing → DEDUCT 1 point from overall score + NOTE IT
- System Design missing/vague → DEDUCT 1.5 points from overall score + NOTE IT

The evaluation is ONLY valid if the student followed the 5-section structure. Missing sections = significant penalties.

════════════════════════════════════════════════
PHASE 1: CLASSIFICATION
════════════════════════════════════════════════

MINDSET: Start by LOOKING FOR A MATCH. Only reject if you are certain it does not fit.
When in doubt → MATCH. A false negative (rejecting valid work) is worse than a false positive.

---

### ⚠️ BEFORE YOU CLASSIFY — CHECK IF DATA EXISTS

Look at the student submission below (Section 01 Problem Statement + Section 02 Use Cases).

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
If problem_match = false → score everything 0, hire_decision = NO, but provide ENCOURAGING feedback acknowledging their effort and inviting them to resubmit with one of the 7 approved problems.
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

Remarks:
{remarks}

GitHub: {github_url}
Deployed: {deployed_url}
Video: {video_url}

Remarks may provide additional context or clarifications — consider them in your evaluation for completeness and insight.

---

---

### Evaluation Criteria (Score 0–10 each — BE VERY STRICT):

1. **Section 01 - Problem Understanding** — Did they clearly describe the problem, pain points, and success criteria? (0–10)
2. **Section 02 - Cases & Logic Constraints** — Did they map important cases, edge cases, and constraints? (0–10)
3. **Section 03 - AI/ML Usage (MOST IMPORTANT)** — Did they meaningfully use AI, showing HOW and WHY they chose specific models? (0–10)
4. **Section 03 - Technical Architecture** — Is the design logical, well-explained, with data flow clarity? (0–10)
5. **Practicality & Scalability** — Would this work in the real world with the AI approach chosen? (0–10)
6. **Clarity & Communication** — Are sections 1–5 clear, well-organized, and easy to follow? (0–10)

### DELIVERABLE VERIFICATION (MANDATORY):
Check these BEFORE calculating overall score:
- **GitHub Quality** — Repo exists, is PUBLIC, has complete README with setup instructions (0–10)
  - Missing GitHub or private → score = 0, DEDUCT 2 from final overall
  - README missing/incomplete → DEDUCT 1 from final overall
- **Deployment Quality** — Live demo accessible showing AI/automation in action (0–10)
  - Missing deployed URL → DEDUCT 2.5 from final overall
  - If deployed exists but broken → score = 0
- **Video Quality** — Screen recording showing code walkthrough and AI in action (0–10)
  - Missing video (and no deployed demo) → DEDUCT 2.5 from final overall
  - If video exists but doesn't show AI/automation → score = 2

### Scoring Rules — STRICT INTERPRETATION:
- Base Overall Score = (avg of criteria 1–6). 
- THEN apply deliverable deductions (GitHub, Deployment, Video penalties above).
- Final Overall Score = Base Score - Total Deductions (cannot go below 0).
- Hire Decision: YES only if Final Overall Score ≥ 7.0 AND all three deliverables (GitHub, Deployed, Video) are present AND GitHub ≥ 6.
- If ANY critical deliverable is missing → Hire Decision = NO (or MAYBE at best).
- Example: Student scores 8.5 on criteria but missing video → 8.5 - 2.5 = 6.0 → NO (below 7.0 threshold).

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
      "cases_constraints": <0.0–10.0>,
      "ai_usage": <0.0–10.0>,
      "technical_architecture": <0.0–10.0>,
      "practicality": <0.0–10.0>,
      "clarity": <0.0–10.0>,
      "github_quality": <0.0–10.0>,
      "deployment_quality": <0.0–10.0>,
      "video_quality": <0.0–10.0>
    }},
    "deliverable_penalties": "Explicitly state each deduction: 'GitHub missing: -2 pts', 'Video missing: -2.5 pts', etc. Total deductions applied.",
    "strengths": [
      "Specific, human observation about their submission",
      "Another genuine strength with reference to their work",
      "A third honest point"
    ],
    "weaknesses": [
      "Honest, specific gap referencing their actual submission",
      "Another concrete weakness",
      "Missing/incomplete section: [specify which section and why it matters]"
    ],
    "ai_feedback": "Conversational paragraph. For IN-SCOPE COMPLETE: 'I noticed you used [specific AI/model]. Your approach to [specific aspect] shows [observation].' For INCOMPLETE: 'You submitted [sections 1-3], but [missing sections]. This makes it hard to fully evaluate your work. To be considered, please add [specific missing section].' For OUT-OF-SCOPE: Acknowledge effort, explain it's outside the 7 problems, encourage resubmission with one of the 7.",
    "improvement_suggestions": [
      "Concrete, actionable suggestion 1 (e.g., 'Add a video showing the AI in action')",
      "Concrete, actionable suggestion 2"
    ],
    "hire_decision": "YES / NO / MAYBE",
    "reason": "Include ALL reasons: (1) Problem match/scope, (2) Quality of sections submitted, (3) Status of all 5 deliverables, (4) Missing penalties applied. Example: 'YES: Problem #3 match, strong section 1-3, GitHub complete with good README, deployed demo works, video shows AI logic. NO: Scores well (8.2) but missing deployed URL and video (-5 pts total) → final 3.2, below threshold. Resubmit with working demo and walkthrough video. MAYBE: Waiting for clarification on [section].' Be specific about missing components and next steps."
  }}
}}

---

### TONE RULES

✅ Write like a human who actually read their submission
✅ Reference specific things they wrote — not generic observations  
✅ "I noticed you used Claude to classify X" not "AI was used effectively"
✅ Weaknesses should feel like coaching, not rejection
✅ FOR OUT-OF-SCOPE SUBMISSIONS: Be encouraging! Acknowledge the effort, the problem-solving mindset, and the technical approach. Then friendly guide them to one of the 7 problems.
✅ FOR INCOMPLETE SUBMISSIONS: Be CLEAR about what's missing and why it matters. Offer a straightforward path to improvement.
❌ Never write robotic filler like "The solution demonstrates architectural competence"
❌ Never make submissions feel punished unfairly — deductions are fair and explained

Now execute Phase 1 (look for a match first), then Phase 2 (human tone). Return ONLY the JSON.
"""

EXPECTED_COLUMNS = [
    "Student Name",
    "Problem Statement",
    "Use Cases",
    "System Design",
    "Remarks",
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

def validate_url(url, type_):
    if not url or url == "Not provided":
        return 0  # Missing
    try:
        if type_ == "github":
            if "github.com" not in url:
                return 0
            # Extract owner/repo from URL
            match = re.search(r'github\.com/([^/]+)/([^/]+)', url)
            if not match:
                return 0
            owner, repo = match.groups()
            # Try main branch first, then master
            readme_urls = [
                f"https://raw.githubusercontent.com/{owner}/{repo}/main/README.md",
                f"https://raw.githubusercontent.com/{owner}/{repo}/master/README.md"
            ]
            readme_content = None
            for readme_url in readme_urls:
                response = requests.get(readme_url, timeout=10)
                if response.status_code == 200:
                    readme_content = response.text
                    break
            if not readme_content:
                return 2  # Repo exists but no README
            # Evaluate README quality
            score = evaluate_readme(readme_content)
            return score
        elif type_ == "deployed":
            response = requests.get(url, timeout=10)
            return 10 if response.status_code == 200 else 0
        elif type_ == "video":
            # Basic check for YouTube or Drive
            if "youtube.com" in url or "youtu.be" in url:
                # Extract video ID
                video_id = None
                if "youtube.com/watch?v=" in url:
                    video_id = url.split("v=")[1].split("&")[0]
                elif "youtu.be/" in url:
                    video_id = url.split("/")[-1]
                if video_id:
                    # Check if video is accessible (public)
                    embed_url = f"https://www.youtube.com/embed/{video_id}"
                    response = requests.head(embed_url, timeout=5)
                    if response.status_code == 200:
                        # Optionally, get title/description if API key available
                        api_key = os.environ.get("YOUTUBE_API_KEY")
                        if api_key:
                            details = get_youtube_details(video_id, api_key)
                            score = evaluate_video_details(details)
                            return score
                        else:
                            return 8  # Accessible, assume good
                    else:
                        return 0  # Private or unavailable
                return 0
            elif "drive.google.com" in url:
                response = requests.head(url, timeout=5)
                return 10 if response.status_code == 200 else 0
            return 0
    except:
        return 0
    return 10  # Valid

def evaluate_readme(content):
    score = 0
    content_lower = content.lower()
    # Check for basic sections
    sections = ["installation", "usage", "setup", "getting started", "features", "contributing", "license"]
    found_sections = sum(1 for section in sections if section in content_lower)
    score += min(found_sections * 2, 6)  # Up to 6 points for sections
    # Check length (at least 200 chars for decent README)
    if len(content) > 200:
        score += 2
    # Check for code blocks or links
    if "```" in content or "[" in content:
        score += 2
    return min(score, 10)  # Cap at 10

def get_youtube_details(video_id, api_key):
    url = f"https://www.googleapis.com/youtube/v3/videos?id={video_id}&key={api_key}&part=snippet"
    response = requests.get(url, timeout=10)
    if response.status_code == 200:
        data = response.json()
        if data['items']:
            snippet = data['items'][0]['snippet']
            return {
                'title': snippet.get('title', ''),
                'description': snippet.get('description', ''),
                'tags': snippet.get('tags', [])
            }
    return {}

def evaluate_video_details(details):
    score = 5  # Base for accessible video
    if not details:
        return score
    title = details.get('title', '').lower()
    description = details.get('description', '').lower()
    # Check if title/description mentions assignment, demo, etc.
    keywords = ['assignment', 'demo', 'walkthrough', 'explanation', 'project']
    if any(k in title or k in description for k in keywords):
        score += 2
    # Check description length (indicates detail)
    if len(description) > 100:
        score += 2
    # Check for tags related to tech
    tags = details.get('tags', [])
    tech_tags = ['python', 'javascript', 'ai', 'ml', 'web', 'app']
    if any(t.lower() in tech_tags for t in tags):
        score += 1
    return min(score, 10)

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

    remarks = data.get("remarks") or "Not provided"

    github_score = ""
    deployed_score = ""
    video_score = ""

    prompt = EVALUATION_PROMPT.format(
        problem_statement=problem_statement,
        use_cases=use_cases,
        system_design=system_design,
        remarks=remarks,
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
                    r.get("Problem Match", "N/A"),
                    r.get("Matched Problem", ""),
                    r.get("Problem Matching Reason", ""),
                    r.get("Problem Understanding", ""),
                    r.get("Cases & Constraints", ""),
                    r.get("AI Usage", ""),
                    r.get("Technical Architecture", ""),
                    r.get("Practicality", ""),
                    r.get("Clarity", ""),
                    r.get("GitHub Quality", ""),
                    r.get("Deployment Quality", ""),
                    r.get("Video Quality", ""),
                    r.get("Deliverable Penalties", ""),
                    r.get("AI Overall Score", ""),
                    r.get("AI Hire Decision", ""),
                    r.get("Strengths", ""),
                    r.get("Weaknesses", ""),
                    r.get("AI Feedback", ""),
                    r.get("Hire Reason", ""),
                    r.get("Human overall score", ""),
                    r.get("Final Decision", ""),
                ])
        except Exception:
            existing_rows = []
            existing_ids = set()

        headers = [
            # Student Identification
            "Name",
            "NIAT ID",
            # Problem Classification (first validation)
            "Problem Match",
            "Matched Problem",
            "Problem Matching Reason",
            # Detailed AI Evaluation Scores
            "Problem Understanding",
            "Cases & Constraints",
            "AI Usage",
            "Technical Architecture",
            "Practicality",
            "Clarity",
            "GitHub Quality",
            "Deployment Quality",
            "Video Quality",
            "Deliverable Penalties",
            # AI Overall Decision
            "AI Overall Score",
            "AI Hire Decision",
            # AI's Detailed Feedback
            "Strengths",
            "Weaknesses",
            "AI Feedback",
            "Hire Reason",
            # Human Review (Final Decision)
            "Human overall score",
            "Final Decision",
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
                deliverable_penalties = evaluation.get("deliverable_penalties", "")
                
                # Format for Sheets
                problem_match_text = "Yes" if problem_match else "No"
                
                rows.append([
                    name,
                    niat_id,
                    problem_match_text,
                    str(matched_problem_number) if matched_problem_number else "",
                    classification_reasoning,
                    s.get("problem_understanding", ""),
                    s.get("cases_constraints", ""),
                    s.get("ai_usage", ""),
                    s.get("technical_architecture", ""),
                    s.get("practicality", ""),
                    s.get("clarity", ""),
                    s.get("github_quality", ""),
                    s.get("deployment_quality", ""),
                    s.get("video_quality", ""),
                    deliverable_penalties,
                    overall_score,
                    hire_decision,
                    " | ".join(strengths),
                    " | ".join(weaknesses),
                    ai_feedback,
                    reason,
                    "",
                    hire_decision,
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
