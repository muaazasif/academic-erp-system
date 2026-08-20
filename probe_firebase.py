import urllib.request
import json

base_url = "https://muaazasif-4cf8e-default-rtdb.firebaseio.com"
paths = ["", "students", "admin", "attendance", "users", "assignments", "quizzes", "marks"]

for path in paths:
    url = f"{base_url}/{path}.json" if path else f"{base_url}/.json"
    print(f"Checking path: {url}")
    try:
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response:
            status = response.status
            content = response.read().decode('utf-8')
            data = json.loads(content)
            if data is not None:
                print(f"  ✅ SUCCESS: {url} -> found keys/data! keys: {list(data.keys()) if isinstance(data, dict) else 'non-dict'}")
                with open(f"firebase_{path or 'root'}.json", "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            else:
                print(f"  ℹ️ SUCCESS but Empty (None)")
    except Exception as e:
        print(f"  ❌ Failed: {e}")
