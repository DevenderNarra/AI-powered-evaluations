import os
import io
import json
from openai import OpenAI, AuthenticationError, APIConnectionError, APIStatusError
import pandas as pd
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, send_file
from dotenv import load_dotenv
load_dotenv()


app = Flask(__name__)

EVALUATION_PROMPT = """
You are a Senior AI Product Engineer and Hiring Evaluator.

Your task is to evaluate a student's project submission based on real-world product thinking, AI usage, and engineering quality.

---

### 📌 Evaluation Context

The student was given the following instructions:

- Pick one problem statement
- Understand the problem deeply before coding
- Identify users and pain points
- Define success criteria
- Build an AI-powered solution (NLP / ML / Generative AI)
- Ensure the solution is practical, simple, and scalable

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
    # Add one example row
    df.loc[0] = [
        "Jane Doe",
        "Students struggle to find peer tutors easily on campus.",
        "Use case 1: Student searches by subject. Edge case: No tutors available.",
        "Input: search query. Rules: match by subject/availability. Failure: fallback to waitlist.",
        "https://github.com/example/repo",
        "https://myapp.vercel.app",
        "https://youtu.be/example",
    ]
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


if __name__ == "__main__":
    app.run(debug=True, port=5000)
