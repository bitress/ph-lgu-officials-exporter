import os
import sys
import json
import time
import threading
import urllib.request

WORKSPACE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if WORKSPACE_DIR not in sys.path:
    sys.path.insert(0, WORKSPACE_DIR)

from app import run_web

PORT = 8899

# Start server in thread
server_thread = threading.Thread(
    target=run_web,
    kwargs={"port": PORT, "data_path": os.path.join(WORKSPACE_DIR, "officials_2025_2028.xlsx")},
    daemon=True
)
server_thread.start()

# Wait for server to bind and be ready
for _ in range(60):
    try:
        with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/meta", timeout=1) as res:
            if res.status == 200:
                break
    except Exception:
        time.sleep(0.1)

# Test 1: GET / (HTML template)
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/") as res:
    html = res.read().decode("utf-8")
    assert "Philippine" in html
    assert "statMayors" in html
    assert "lguTypeSelect" in html
    print("Test 1 (GET /): OK")

# Test 2: GET /api/meta
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/meta") as res:
    meta = json.loads(res.read().decode("utf-8"))
    assert len(meta["regions"]) == 18
    print("Test 2 (GET /api/meta): OK")

# Test 3: GET /api/stats
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/stats") as res:
    stats = json.loads(res.read().decode("utf-8"))
    assert stats["total_lgus"] == 1633
    assert stats["total_municipalities"] == 1484
    assert stats["total_cities"] == 149
    assert stats["addresses_count"] == 1633
    print("Test 3 (GET /api/stats): OK")

# Test 4: GET /api/records?region=NCR&lgu_type=municipality
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/records?region=NCR&lgu_type=municipality") as res:
    recs = json.loads(res.read().decode("utf-8"))
    assert len(recs) == 1
    assert recs[0]["Town"] == "Pateros"
    assert recs[0]["LGU Type"] == "Municipality"
    print("Test 4 (GET /api/records?region=NCR&lgu_type=municipality): OK")

# Test 5: GET /api/records?region=NCR&lgu_type=city
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/records?region=NCR&lgu_type=city") as res:
    recs = json.loads(res.read().decode("utf-8"))
    assert len(recs) == 16
    assert all(r["LGU Type"] == "City" for r in recs)
    print("Test 5 (GET /api/records?region=NCR&lgu_type=city): OK")

# Test 6: GET /api/records?region=NCR (all LGUs)
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/records?region=NCR") as res:
    recs = json.loads(res.read().decode("utf-8"))
    assert len(recs) == 17
    print("Test 6 (GET /api/records?region=NCR [all]): OK")

# Test 7: GET /api/export?format=xlsx&region=NCR&lgu_type=city
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/export?format=xlsx&region=NCR&lgu_type=city") as res:
    data = res.read()
    assert len(data) > 1000
    assert res.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    print("Test 7 (GET /api/export xlsx): OK")

# Test 8: GET /api/export?format=csv&lgu_type=city&include_lgu_type=1
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/export?format=csv&lgu_type=city&include_lgu_type=1") as res:
    csv_text = res.read().decode("utf-8-sig")
    lines = csv_text.strip().splitlines()
    assert len(lines) == 150  # 1 header + 149 cities
    assert "LGU Type" in lines[0]
    print("Test 8 (GET /api/export csv with LGU Type): OK")

# Test 9: GET /verify (HTML page)
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/verify") as res:
    html = res.read().decode("utf-8")
    assert "Contact Verification" in html
    assert "verificationMap" in html
    assert "accordionContainer" in html
    print("Test 9 (GET /verify): OK")

# Test 10: GET /api/verification
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/verification") as res:
    vmap = json.loads(res.read().decode("utf-8"))
    assert isinstance(vmap, dict)
    print("Test 10 (GET /api/verification): OK")

# Test 11: POST /api/verification (verify an LGU)
post_data = json.dumps({"province": "Abra", "town": "Bangued (capital)", "verified": True}).encode("utf-8")
req = urllib.request.Request(
    f"http://127.0.0.1:{PORT}/api/verification",
    data=post_data,
    headers={"Content-Type": "application/json"}
)
with urllib.request.urlopen(req) as res:
    resp = json.loads(res.read().decode("utf-8"))
    assert resp["success"] is True
    assert resp["timestamp"] is not None
    print("Test 11 (POST /api/verification verify): OK")

# Test 12: GET /api/export with verified_only=1 and include_verified_col=1
with urllib.request.urlopen(f"http://127.0.0.1:{PORT}/api/export?format=json&verified_only=1&include_verified_col=1") as res:
    exp = json.loads(res.read().decode("utf-8"))
    assert len(exp) == 1
    assert exp[0]["Town"] == "Bangued (capital)"
    assert exp[0]["Contact Verified"] == "Yes"
    print("Test 12 (GET /api/export verified_only): OK")

# Test 13: POST /api/verification (unverify to clean up)
unverify_data = json.dumps({"province": "Abra", "town": "Bangued (capital)", "verified": False}).encode("utf-8")
req = urllib.request.Request(
    f"http://127.0.0.1:{PORT}/api/verification",
    data=unverify_data,
    headers={"Content-Type": "application/json"}
)
with urllib.request.urlopen(req) as res:
    resp = json.loads(res.read().decode("utf-8"))
    assert resp["success"] is True
    assert resp["timestamp"] is None
    print("Test 13 (POST /api/verification unverify): OK")

print("\nALL WEB TESTS PASSED!")
