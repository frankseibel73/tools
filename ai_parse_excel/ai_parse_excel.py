import pandas as pd
import requests
import json
import time

LM_URL = "http://192.168.2.3:1234v1/chat/completions"
MODEL = "deepseek-coder-v2-lite-instruct"

df = pd.read_excel("/data/time_entries.xlsx")

def analyze_entry(text):
    prompt = f"""
Analyze this developer time entry.

Return ONLY JSON with these fields:
category
activity
component
work_type
customer_related
summary

Entry:
{text}
"""

    response = requests.post(
        LM_URL,
        json={
            "model": MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0
        }
    )

    content = response.json()["choices"][0]["message"]["content"]

    try:
        return json.loads(content)
    except:
        return {
            "category": "",
            "activity": "",
            "component": "",
            "work_type": "",
            "customer_related": "",
            "summary": ""
        }
```python ai_parse_excel/ai_parse_excel.py
import pandas as pd
import requests
import json
import time

LM_URL = "http://192.168.2.3:1234v1/chat/completions"
MODEL = "deepseek-coder-v2-lite-instruct"

df = pd.read_excel("/data/time_entries.xlsx")

def analyze_entry(text):
    prompt = f"""
Analyze this developer time entry.

Return ONLY JSON with these fields:
category
activity
component
work_type
customer_related
summary

Entry:
{text}
"""

    response = requests.post(
        LM_URL,
        json={
            "model": MODEL,
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0
        }
    )

    content = response.json()["choices"][0]["message"]["content"]

    try:
        return json.loads(content)
    except:
        return {
            "category": "",
            "activity": "",
            "component": "",
            "work_type": "",
            "customer_related": "",
            "summary": ""
        }

df = pd.read_excel("/data/time_entries.xlsx")
results = []

for index, row in df.iterrows():
    entry = ", ".join([str(row[col]) for col in df.columns])
    results.append(analyze_entry(entry))
    time.sleep(0.2)  # prevents overloading the model

analysis_df = pd.DataFrame(results)

df = pd.concat([df, analysis_df], axis=1)

df.to_excel("/data/time_entries_analyzed.xlsx", index=False)