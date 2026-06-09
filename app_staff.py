import os
import io
import jwt
import datetime
import pandas as pd
from flask import jsonify, request
from supabase import create_client

SUPABASE_STAFF_URL = os.environ.get("SUPABASE_STAFF_URL")
SUPABASE_STAFF_KEY = os.environ.get("SUPABASE_STAFF_KEY")
JWT_SECRET = os.environ.get("JWT_SECRET")

supabase_staff = create_client(SUPABASE_STAFF_URL, SUPABASE_STAFF_KEY)

def register_staff_routes(app):

    @app.route("/health_staff")
    def health_staff():
        return jsonify({"status": "ok", "service": "staff-dashboard"})

    @app.route("/staff/upload/appointments_json", methods=["POST"])
    def upload_appointments_json():
        try:
            data = request.get_json()
            records = data.get("records", [])
            if not records:
                return jsonify({"error": "データがありません"}), 400
            for r in records:
                r["target_month"] = r["acquired_date"][:7] + "-01" if r.get("acquired_date") else None
            supabase_staff.table("appointments").upsert(records).execute()
            return jsonify({"status": "ok", "count": len(records)})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/staff/upload/productivity_json", methods=["POST"])
    def upload_productivity_json():
        try:
            data = request.get_json()
            records = data.get("records", [])
            if not records:
                return jsonify({"error": "データがありません"}), 400
            for r in records:
                r["target_month"] = r["call_date"][:7] + "-01" if r.get("call_date") else None
            supabase_staff.table("productivity").upsert(records).execute()
            return jsonify({"status": "ok", "count": len(records)})
        except Exception as e:
            return jsonify({"error": str(e)}), 500

    @app.route("/staff/upload/attendance_json", methods=["POST"])
    def upload_attendance_json():
        try:
            data = request.get_json()
            records = data.get("records", [])
            if not records:
                return jsonify({"error": "データがありません"}), 400
            for r in records:
                r["target_month"] = r["work_date"][:7] + "-01" if r.get("work_date") else None
            supabase_staff.table("attendance").upsert(records).execute()
            return jsonify({"status": "ok", "count": len(records)})
        except Exception as e:
            return jsonify({"error": str(e)}), 500
