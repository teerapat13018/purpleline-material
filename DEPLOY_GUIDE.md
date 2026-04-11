# 🚀 คู่มือ Deploy ระบบเบิกวัสดุ Purple Line บน GitHub + Streamlit Cloud

## โครงสร้างไฟล์ที่ต้อง Push ขึ้น GitHub

```
purpleline-material/
├── app.py
├── requirements.txt
├── Form เบิกของ TAM 19032026.xlsx
├── signature.png
├── .gitignore                    ← ✅ สร้างแล้ว
└── .streamlit/
    └── secrets.toml              ← ⚠️ อย่า push! (.gitignore บล็อกไว้แล้ว)
```

---

## ขั้นตอนที่ 1 — สร้าง GitHub Repository

1. ไปที่ https://github.com/new
2. ตั้งชื่อ repo เช่น `purpleline-material`
3. เลือก **Private** (แนะนำ เพราะมี business logic)
4. กด **Create repository**

---

## ขั้นตอนที่ 2 — Push โค้ดขึ้น GitHub

เปิด Terminal แล้วรันคำสั่งนี้ใน folder `purpleline-material/`:

```bash
cd path/to/purpleline-material

git init
git add app.py requirements.txt "Form เบิกของ TAM 19032026.xlsx" signature.png .gitignore
git commit -m "Initial commit: Purple Line material requisition app"

git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/purpleline-material.git
git push -u origin main
```

> ⚠️ **อย่าทำ `git add .`** โดยตรง เพราะอาจติด `secrets.toml` ไปด้วย
> ให้ add ทีละไฟล์ตามด้านบนแทน

---

## ขั้นตอนที่ 3 — Deploy บน Streamlit Community Cloud

1. ไปที่ https://share.streamlit.io
2. Sign in ด้วย GitHub account เดียวกัน
3. กด **New app**
4. เลือก:
   - **Repository**: `YOUR_USERNAME/purpleline-material`
   - **Branch**: `main`
   - **Main file path**: `app.py`
5. กด **Advanced settings** → **Secrets**
6. วาง JSON ของ Service Account ในรูปแบบนี้:

```toml
[gcp_service_account]
type = "service_account"
project_id = "YOUR_PROJECT_ID"
private_key_id = "YOUR_PRIVATE_KEY_ID"
private_key = "-----BEGIN RSA PRIVATE KEY-----\nXXXXX\n-----END RSA PRIVATE KEY-----\n"
client_email = "YOUR_SA@YOUR_PROJECT.iam.gserviceaccount.com"
client_id = "YOUR_CLIENT_ID"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/YOUR_SA%40YOUR_PROJECT.iam.gserviceaccount.com"
```

7. กด **Deploy!**

---

## ขั้นตอนที่ 4 — ตรวจสอบ Google Sheets Permission

ตรวจสอบว่า Service Account email ของคุณได้รับสิทธิ์ **Editor** ใน Google Sheet แล้ว:

1. เปิด Google Sheet (ID: `145x1TXKeU8vj_xYyaJz30PtEqitCrsOx`)
2. กด Share → เพิ่ม email ของ Service Account
3. ให้สิทธิ์ **Editor**

---

## เมื่อ Deploy สำเร็จ

- แอปจะได้ URL รูปแบบ: `https://YOUR_USERNAME-purpleline-material-app-XXXXX.streamlit.app`
- แชร์ URL นี้ให้พนักงานใช้งานได้เลย

---

## การอัปเดตโค้ดในอนาคต

```bash
git add app.py  # หรือไฟล์ที่แก้ไข
git commit -m "แก้ไข: ..."
git push
```

Streamlit Cloud จะ **redeploy อัตโนมัติ** ภายใน 1-2 นาที

