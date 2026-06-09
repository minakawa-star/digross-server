import os
import io
import jwt
import datetime
import pandas as pd
from flask import jsonify, request
from supabase import create_client
from flask_cors import CORS

SUPABASE_STAFF_URL = os.environ.get("SUPABASE_STAFF_URL")
SUPABASE_STAFF_KEY = os.environ.get("SUPABASE_STAFF_KEY")
JWT_SECRET = os.environ.get("JWT_SECRET")

supabase_staff = create_client(SUPABASE_STAFF_URL, SUPABASE_STAFF_KEY)

def register_staff_routes(app):
    CORS(app, resources={
        r"/staff/*": {"origins": [
            "https://minakawa-star.github.io",
            "http://localhost:3000"
        ]},
        r"/health_staff": {"origins": "*"}
    })

    @app.route("/health_staff")
    def health_staff():
        return jsonify({"status": "ok", "service": "staff-dashboard"})

    @app.route("/staff/upload/appointments", methods=["POST"])
    def upload_appointments():
        try:
            file = request.files.get("file")
            if not file:
                return jsonify({"error": "ファイルがありません"}), 400

            df = pd.read_excel(io.BytesIO(file.read()))
            df.columns = df.columns.str.strip()

            records = []
            for _, row in df.iterrows():
                staff_id = str(row.get("社員番号", "")).strip()
                if not staff_id:
                    continue

                acquired_date = row.get("取得日")
                cancel_date = str(row.get("キャンセル受付日", "")).strip()
                fb_date = row.get("フィードバック受付日")
                amount = row.get("案件金額", 0)
                fb_amount = row.get("達成金額", 0)

                if pd.notna(acquired_date):
                    target_month = pd.to_datetime(acquired_date).strftime("%Y-%m-01")
                else:
                    target_month = None

                records.append({
                    "staff_id": staff_id,
                    "acquired_date": pd.to_datetime(acquired_date).strftime("%Y-%m-%d") if pd.notna(acquired_date) else None,
                    "amount": int(amount) if pd.notna(amount) else 0,
                    "cancel_date": cancel_date if cancel_date not in ["nan", ""] else None,
                    "fb_date": pd.to_datetime(fb_date).strftime("%Y-%m-%d") if pd.notna(fb_date) else None,
                    "fb_amount": int(fb_amount) if pd.notna(fb_amount) else 0,
                    "target_month": target_month
                })

            if records:
                supabase_staff.table("appointments").upsert(records).execute()

            return jsonify({"status": "ok", "count": len(records)})

        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/staff/upload/productivity", methods=["POST"])
    def upload_productivity():
        try:
            file = request.files.get("file")
            if not file:
                return jsonify({"error": "ファイルがありません"}), 400

            df = pd.read_csv(io.BytesIO(file.read()), encoding="utf-8-sig")
            df.columns = df.columns.str.strip()

            records = []
            for _, row in df.iterrows():
                staff_id = str(row.get("社員番号", "")).strip()
                if not staff_id:
                    continue

                call_date = row.get("日付")
                call_count = row.get("コール件数", 0)

                if pd.notna(call_date):
                    target_month = pd.to_datetime(call_date).strftime("%Y-%m-01")
                else:
                    target_month = None

                records.append({
                    "staff_id": staff_id,
                    "call_date": pd.to_datetime(call_date).strftime("%Y-%m-%d") if pd.notna(call_date) else None,
                    "call_count": int(call_count) if pd.notna(call_count) else 0,
                    "target_month": target_month
                })

            if records:
                supabase_staff.table("productivity").upsert(records).execute()

            return jsonify({"status": "ok", "count": len(records)})

        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/staff/upload/attendance", methods=["POST"])
    def upload_attendance():
        try:
            file = request.files.get("file")
            if not file:
                return jsonify({"error": "ファイルがありません"}), 400

            df = pd.read_csv(io.BytesIO(file.read()), encoding="utf-8-sig")
            df.columns = df.columns.str.strip()

            records = []
            for _, row in df.iterrows():
                staff_id = str(row.get("*従業員ID", "")).strip()
                if not staff_id:
                    continue

                work_date = row.get("*年月日")
                work_time = str(row.get("実労働時間", "00:00")).strip()

                try:
                    h, m = work_time.split(":")
                    work_hours = int(h) + int(m) / 60
                    work_hours = round(work_hours, 3)
                except:
                    work_hours = 0.0

                if pd.notna(work_date):
                    target_month = pd.to_datetime(work_date).strftime("%Y-%m-01")
                else:
                    target_month = None

                records.append({
                    "staff_id": staff_id,
                    "work_date": pd.to_datetime(work_date).strftime("%Y-%m-%d") if pd.notna(work_date) else None,
                    "work_hours": work_hours,
                    "target_month": target_month
                })

            if records:
                supabase_staff.table("attendance").upsert(records).execute()

            return jsonify({"status": "ok", "count": len(records)})

        except Exception as e:
            return jsonify({"error": str(e)}), 500
