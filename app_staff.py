import os
import jwt
import datetime
import bcrypt
from flask import jsonify, request
from supabase import create_client
from openpyxl import load_workbook
from functools import wraps

SUPABASE_STAFF_URL = os.environ.get("SUPABASE_STAFF_URL")
SUPABASE_STAFF_KEY = os.environ.get("SUPABASE_STAFF_KEY")
JWT_SECRET = os.environ.get("JWT_SECRET")

supabase_staff = create_client(SUPABASE_STAFF_URL, SUPABASE_STAFF_KEY)

MASTER_PATH = os.path.join(os.path.dirname(__file__), "スタッフマスター.xlsx")

B_TO_D = {
    "B0000106": "D0000295",
    "B0000107": "D0000326",
    "D0001318": "B0000095"
}

RATE_TABLE = {
    5: {22: (2.30, 2.00), 21: (2.30, 2.00), 20: (2.45, 2.10),
        19: (2.45, 2.10), 18: (2.45, 2.10), 17: (2.60, 2.20),
        16: (2.60, 2.20), 15: (2.60, 2.20), 14: (2.60, 2.20),
        13: (2.75, 2.30), 12: (2.75, 2.30), 11: (2.75, 2.30),
        10: (2.75, 2.30), 9: (2.75, 2.30), 8: (2.95, 2.57),
        7: (2.95, 2.57), 6: (2.95, 2.57), 5: (2.95, 2.57)},
    4: {22: (2.30, 2.00), 21: (2.30, 2.00), 20: (2.45, 2.10),
        19: (2.45, 2.10), 18: (2.45, 2.10), 17: (2.60, 2.20),
        16: (2.60, 2.20), 15: (2.60, 2.20), 14: (2.60, 2.20),
        13: (2.75, 2.30), 12: (2.75, 2.30), 11: (2.75, 2.30),
        10: (2.75, 2.30), 9: (2.75, 2.30), 8: (2.95, 2.57),
        7: (2.95, 2.57), 6: (2.95, 2.57), 5: (2.95, 2.57)},
    3: {22: (2.30, 2.00), 21: (2.30, 2.00), 20: (2.45, 2.10),
        19: (2.45, 2.10), 18: (2.45, 2.10), 17: (2.60, 2.20),
        16: (2.60, 2.20), 15: (2.60, 2.20), 14: (2.60, 2.20),
        13: (2.75, 2.30), 12: (2.75, 2.30), 11: (2.75, 2.30),
        10: (2.75, 2.30), 9: (2.75, 2.30), 8: (2.95, 2.57),
        7: (2.95, 2.57), 6: (2.95, 2.57), 5: (2.95, 2.57)},
}

def load_staff_master():
    wb = load_workbook(MASTER_PATH, data_only=True)
    ws1 = wb["スタッフマスター"]
    ws2 = wb["時給マスター"]

    master = {}
    for row in ws1.iter_rows(min_row=2, values_only=True):
        staff_id, name, site, rank = row[0], row[1], row[2], row[3]
        if not staff_id or not rank:
            continue
        sid = str(staff_id).strip()
        sid = B_TO_D.get(sid, sid)
        master[sid] = {
            "name": name, "site": site, "rank": rank,
            "hourly_wage": 0, "mgmt_fee": 0,
            "work_pattern": 5, "monthly_salary": None
        }

    for row in ws2.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        sid = str(row[0]).strip()
        sid = B_TO_D.get(sid, sid)
        wage = row[2]
        note = str(row[3]) if row[3] else ""
        if sid not in master:
            continue
        if "月給" in note:
            master[sid]["monthly_salary"] = int(str(wage).replace(",", "").strip()) if wage else 0
        else:
            master[sid]["hourly_wage"] = int(str(wage).replace(",", "").strip()) if wage else 0
        master[sid]["mgmt_fee"] = 3030 if "管理料" in note else 0

    return master

def jwt_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("Authorization", "").replace("Bearer ", "")
        if not token:
            return jsonify({"error": "認証が必要です"}), 401
        try:
            payload = jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
            request.staff_id = payload.get("staff_id")
            request.role = payload.get("role")
        except jwt.ExpiredSignatureError:
            return jsonify({"error": "トークンが期限切れです"}), 401
        except jwt.InvalidTokenError:
            return jsonify({"error": "無効なトークンです"}), 401
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        token = request.headers.get("Authorization", "").replace("Bearer ", "")
        if not token:
            return jsonify({"error": "認証が必要です"}), 401
        try:
            payload = jwt.decode(token, JWT_SECRET, algorithms=["HS256"])
            if payload.get("role") != "admin":
                return jsonify({"error": "管理者権限が必要です"}), 403
            request.staff_id = payload.get("staff_id")
            request.role = payload.get("role")
        except jwt.InvalidTokenError:
            return jsonify({"error": "無効なトークンです"}), 401
        return f(*args, **kwargs)
    return decorated

def register_staff_routes(app):

    @app.route("/health_staff")
    def health_staff():
        return jsonify({"status": "ok", "service": "staff-dashboard"})

    @app.route("/staff/login", methods=["POST"])
    def staff_login():
        try:
            data = request.get_json()
            login_id = data.get("login_id", "").strip()
            password = data.get("password", "").strip()
            if not login_id or not password:
                return jsonify({"error": "IDとパスワードを入力してください"}), 400
            res = supabase_staff.table("staff_master")\
                .select("*").eq("login_id", login_id).execute()
            if not res.data:
                return jsonify({"error": "IDまたはパスワードが間違っています"}), 401
            staff = res.data[0]
            if not bcrypt.checkpw(password.encode(), staff["password_hash"].encode()):
                return jsonify({"error": "IDまたはパスワードが間違っています"}), 401
            token = jwt.encode({
                "staff_id": staff["staff_id"],
                "role": staff["role"],
                "exp": datetime.datetime.utcnow() + datetime.timedelta(days=30)
            }, JWT_SECRET, algorithm="HS256")
            return jsonify({
                "status": "ok",
                "token": token,
                "staff_id": staff["staff_id"],
                "role": staff["role"],
                "name": staff["staff_name"]
            })
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/register", methods=["POST"])
    @admin_required
    def staff_register():
        try:
            data = request.get_json()
            staff_id = data.get("staff_id", "").strip()
            staff_name = data.get("staff_name", "").strip()
            login_id = data.get("login_id", "").strip()
            password = data.get("password", "").strip()
            role = data.get("role", "staff")
            if not all([staff_id, staff_name, login_id, password]):
                return jsonify({"error": "必須項目が不足しています"}), 400
            password_hash = bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()
            supabase_staff.table("staff_master").upsert({
                "staff_id": staff_id,
                "staff_name": staff_name,
                "login_id": login_id,
                "password_hash": password_hash,
                "role": role
            }).execute()
            return jsonify({"status": "ok", "staff_id": staff_id})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/me")
    @jwt_required
    def staff_me():
        try:
            res = supabase_staff.table("staff_master")\
                .select("staff_id,staff_name,role")\
                .eq("staff_id", request.staff_id).execute()
            if not res.data:
                return jsonify({"error": "スタッフが見つかりません"}), 404
            return jsonify({"status": "ok", "data": res.data[0]})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/staff/debug_master")
    def debug_master():
        try:
            master = load_staff_master()
            targets = ["B0000002", "B0000032", "D0000295", "D0000326", "D0001221", "D0001316"]
            result = {k: v for k, v in master.items() if k in targets}
            return jsonify(result)
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/target", methods=["GET", "POST"])
    @jwt_required
    def staff_target():
        if request.method == "GET":
            try:
                staff_id = request.args.get("staff_id")
                month = request.args.get("month")
                if not staff_id or not month:
                    return jsonify({"error": "staff_id, monthが必要です"}), 400

                if request.role != "admin" and staff_id != request.staff_id:
                    return jsonify({"error": "権限がありません"}), 403

                target_month = month + "-01"
                res = supabase_staff.table("monthly_targets")\
                    .select("*").eq("staff_id", staff_id).eq("target_month", target_month).execute()

                if res.data:
                    return jsonify({"status": "ok", "data": res.data[0]})
                else:
                    return jsonify({"status": "ok", "data": {
                        "staff_id": staff_id, "target_month": target_month,
                        "planned_work_days": 0, "is_confirmed": False, "confirmed_work_days": None
                    }})
            except Exception as e:
                import traceback
                return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

        else:
            try:
                data = request.get_json()
                staff_id = data.get("staff_id")
                month = data.get("month")
                planned_work_days = data.get("planned_work_days")

                if not staff_id or not month:
                    return jsonify({"error": "staff_id, monthが必要です"}), 400

                if request.role != "admin" and staff_id != request.staff_id:
                    return jsonify({"error": "権限がありません"}), 403

                target_month = month + "-01"

                if request.role != "admin":
                    existing = supabase_staff.table("monthly_targets")\
                        .select("is_confirmed").eq("staff_id", staff_id).eq("target_month", target_month).execute()
                    if existing.data and existing.data[0]["is_confirmed"]:
                        return jsonify({"error": "確定済みのため編集できません"}), 403

                supabase_staff.table("monthly_targets").upsert({
                    "staff_id": staff_id,
                    "target_month": target_month,
                    "planned_work_days": int(planned_work_days)
                }, on_conflict="staff_id,target_month").execute()

                return jsonify({"status": "ok"})
            except Exception as e:
                import traceback
                return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/confirm_month", methods=["POST"])
    @admin_required
    def confirm_month():
        try:
            data = request.get_json()
            month = data.get("month")
            if not month:
                return jsonify({"error": "monthが必要です"}), 400

            target_month = month + "-01"
            master = load_staff_master()

            att_res = supabase_staff.table("attendance")\
                .select("*").eq("target_month", target_month).execute()

            work_days_map = {}
            for row in att_res.data:
                sid = B_TO_D.get(row["staff_id"], row["staff_id"])
                if (row.get("work_hours") or 0) > 0:
                    work_days_map[sid] = work_days_map.get(sid, 0) + 1

            for sid in master.keys():
                days = work_days_map.get(sid, 0)
                supabase_staff.table("monthly_targets").upsert({
                    "staff_id": sid,
                    "target_month": target_month,
                    "is_confirmed": True,
                    "confirmed_work_days": days
                }, on_conflict="staff_id,target_month").execute()

            return jsonify({"status": "ok", "message": f"{month}を確定しました"})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/summary")
    def staff_summary():
        try:
            month = request.args.get("month")
            if not month:
                return jsonify({"error": "monthパラメータが必要です"}), 400

            target_month = month + "-01"
            master = load_staff_master()

            apo_res = supabase_staff.table("appointments")\
                .select("*").eq("target_month", target_month).execute()
            apo_rows = apo_res.data

            att_res = supabase_staff.table("attendance")\
                .select("*").eq("target_month", target_month).execute()
            att_rows = att_res.data

            # 出勤予定/確定データ取得
            targets_res = supabase_staff.table("monthly_targets")\
                .select("*").eq("target_month", target_month).execute()
            targets_map = {t["staff_id"]: t for t in targets_res.data}

            results = {}
            for sid, info in master.items():
                results[sid] = {
                    "staff_id": sid,
                    "name": info["name"],
                    "site": info["site"],
                    "rank": info["rank"],
                    "apo_amount": 0,
                    "cxl_amount": 0,
                    "fb_amount": 0,
                    "sales": 0,
                    "work_days": 0,
                    "target_achieve": 0,
                    "target_maintain": 0,
                    "achieve_rate": None,
                    "is_monthly": info["monthly_salary"] is not None,
                    "hourly_wage": info["hourly_wage"],
                    "monthly_salary": info["monthly_salary"],
                    "planned_work_days": 0,
                    "is_confirmed": False
                }

            for row in apo_rows:
                sid = B_TO_D.get(row["staff_id"], row["staff_id"])
                if sid not in results:
                    continue
                cancel = str(row.get("cancel_date") or "")
                if cancel and cancel not in ["None", ""]:
                    results[sid]["cxl_amount"] += row.get("amount", 0)
                else:
                    results[sid]["apo_amount"] += row.get("amount", 0)
                results[sid]["fb_amount"] += row.get("fb_amount", 0)

            for row in att_rows:
                sid = B_TO_D.get(row["staff_id"], row["staff_id"])
                if sid not in results:
                    continue
                if (row.get("work_hours") or 0) > 0:
                    results[sid]["work_days"] += 1

            for sid, r in results.items():
                info = master[sid]
                r["sales"] = r["apo_amount"] - r["cxl_amount"] + r["fb_amount"]

                tgt = targets_map.get(sid)
                is_confirmed = tgt["is_confirmed"] if tgt else False
                r["is_confirmed"] = is_confirmed
                r["planned_work_days"] = tgt["planned_work_days"] if tgt else 0

                if is_confirmed:
                    calc_days = tgt.get("confirmed_work_days")
                    if calc_days is None:
                        calc_days = r["work_days"]
                else:
                    calc_days = tgt["planned_work_days"] if (tgt and tgt["planned_work_days"] > 0) else r["work_days"]

                if info["monthly_salary"] is not None:
                    base = info["monthly_salary"] * 1.15 + 20000
                    r["target_achieve"] = int(base / 0.40)
                    r["target_maintain"] = int(base / 0.45)
                    if r["work_days"] == 0:
                        r["work_days"] = 22
                else:
                    wage = info["hourly_wage"]
                    mgmt = info["mgmt_fee"]
                    pattern = info["work_pattern"]
                    days = calc_days
                    rate_row = RATE_TABLE.get(pattern, {}).get(days)
                    if rate_row:
                        base = wage * 8 + 1000 + mgmt
                        r["target_achieve"] = int(base * days * rate_row[0])
                        r["target_maintain"] = int(base * days * rate_row[1])

                if r["target_achieve"] > 0:
                    r["achieve_rate"] = round(r["sales"] / r["target_achieve"] * 100, 1)

            return jsonify({"status": "ok", "data": list(results.values())})

        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/upload/appointments_json", methods=["POST"])
    def upload_appointments_json():
        try:
            data = request.get_json()
            records = data.get("records", [])
            if not records:
                return jsonify({"error": "データがありません"}), 400
            for r in records:
                r["target_month"] = r["acquired_date"][:7] + "-01" if r.get("acquired_date") else None
            supabase_staff.table("appointments").upsert(records, on_conflict="appointment_id").execute()
            return jsonify({"status": "ok", "count": len(records)})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/upload/productivity_json", methods=["POST"])
    def upload_productivity_json():
        try:
            data = request.get_json()
            records = data.get("records", [])
            if not records:
                return jsonify({"error": "データがありません"}), 400
            for r in records:
                r["target_month"] = r["call_date"][:7] + "-01" if r.get("call_date") else None
            supabase_staff.table("productivity").upsert(records, on_conflict="staff_id,call_date").execute()
            return jsonify({"status": "ok", "count": len(records)})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500

    @app.route("/staff/upload/attendance_json", methods=["POST"])
    def upload_attendance_json():
        try:
            data = request.get_json()
            records = data.get("records", [])
            if not records:
                return jsonify({"error": "データがありません"}), 400
            for r in records:
                r["target_month"] = r["work_date"][:7] + "-01" if r.get("work_date") else None
            supabase_staff.table("attendance").upsert(records, on_conflict="staff_id,work_date").execute()
            return jsonify({"status": "ok", "count": len(records)})
        except Exception as e:
            import traceback
            return jsonify({"error": str(e), "trace": traceback.format_exc()}), 500
