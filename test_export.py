# app.py  ->  streamlit run app.py
import csv
import io
import re
import zipfile
from datetime import datetime

import pandas as pd
import requests
import streamlit as st
import xlsxwriter
from lxml import etree

# ======================= إعدادات ثابتة =======================
URL = "https://ejada.sa"
DB = "ejada"
MODEL = "crm.lead"
PORTFOLIO_LABEL = "Portfolio Type"
BASE_DOMAIN = [["deleted_user", "=", False]]      # استبعاد المحذوف
BATCH = 50000                                      # لو السيرفر اتأخر نزّله لـ 20000
TEST_LIMIT = 5000
EXCEL_MAX = 1_048_575                              # حد صفوف الإكسيل (بدون الهيدر)
HIDE_PREFIX = ("message_", "activity_", "website_message", "has_message", "__")
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NEWLINES = re.compile(r"\s*[\r\n]+\s*")

st.set_page_config(page_title="سحب بيانات Claims", page_icon="📥", layout="centered")
st.markdown(
    """
    <style>
    .block-container {max-width: 860px; padding-top: 2rem;}
    .stApp, [data-testid="stSidebar"] {direction: rtl;}
    [data-testid="stDataFrame"], [data-testid="stFileUploader"] {direction: ltr;}
    h1, h2, h3 {text-align: right;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ======================= Odoo JSON-RPC =======================
def rpc(service, method, *args):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {"service": service, "method": method, "args": args},
        "id": 1,
    }
    r = requests.post(f"{URL}/jsonrpc", json=payload, timeout=900)
    r.raise_for_status()
    res = r.json()
    if "error" in res:
        err = res["error"]
        raise Exception(err.get("data", {}).get("message") or err.get("message") or str(err))
    return res["result"]


def call(a, model, method, args, kwargs=None):
    return rpc("object", "execute_kw", DB, a["uid"], a["pwd"], model, method, args, kwargs or {})


# ======================= تحميل البيانات الأساسية =======================
def path_label(path, labels):
    parts = path.split("/")
    first = labels.get(parts[0], parts[0])
    return first if len(parts) == 1 else first + "/" + "/".join(parts[1:])


def load_everything(a):
    meta = call(a, MODEL, "fields_get", [], {"attributes": ["string", "type"]})

    portfolio = next((n for n, m in meta.items()
                      if m.get("string", "").strip().lower() == PORTFOLIO_LABEL.lower()), None)
    if not portfolio:
        raise Exception(f"مش لاقي حقل '{PORTFOLIO_LABEL}'، كلّم المسؤول عن الأداة.")

    labels = {n: m.get("string", n) for n, m in meta.items()}
    fields = {n: labels[n] for n, m in meta.items()
              if m.get("type") != "binary" and not n.startswith(HIDE_PREFIX)}
    count = {}
    for v in fields.values():
        count[v] = count.get(v, 0) + 1
    display = {n: (v if count[v] == 1 else f"{v}  [{n}]") for n, v in fields.items()}

    groups = []
    raw = call(a, MODEL, "read_group", [BASE_DOMAIN, [portfolio], [portfolio]],
               {"lazy": False, "limit": 500})
    for g in raw:
        val = g.get(portfolio)
        cnt = g.get("__count", 0)
        if isinstance(val, (list, tuple)):
            key, name = val[0], val[1]
        else:
            key, name = val, (str(val) if val not in (False, None) else "بدون قيمة")
        groups.append({"key": key, "name": name, "count": cnt})
    groups.sort(key=lambda x: x["name"].lower())

    templates = {}
    try:
        exports = call(a, "ir.exports", "search_read", [[["resource", "=", MODEL]]],
                       {"fields": ["name", "export_fields"]})
        for e in exports:
            lines = call(a, "ir.exports.line", "read", [e["export_fields"]], {"fields": ["name"]}) \
                if e["export_fields"] else []
            templates[e["name"]] = [l["name"] for l in lines]
    except Exception:
        templates = {}

    return dict(portfolio=portfolio, labels=labels, display=display, groups=groups,
                templates=templates,
                field_names=sorted(display, key=lambda n: display[n].lower()))


# ======================= السحب + الكتابة في الملف (دفعة بدفعة) =======================
def flat(v):
    """يشيل الأسطر الجديدة من النصوص (عشان مفيش wrap)"""
    if isinstance(v, str) and ("\n" in v or "\r" in v):
        return NEWLINES.sub(" ", v).strip()
    return v


def build_file(a, domain, paths, headers, limit, fmt, progress):
    kw = {"order": "id asc"}
    if limit:
        kw["limit"] = limit
    ids = call(a, MODEL, "search", [domain], kw)
    total = len(ids)
    if not total:
        return None

    is_xlsx = fmt == "Excel"
    buf = io.BytesIO()
    preview, n, truncated = [], 0, False

    if is_xlsx:
        wb = xlsxwriter.Workbook(buf, {"constant_memory": True,
                                       "strings_to_urls": False,
                                       "strings_to_formulas": False})
        ws = wb.add_worksheet("Data")
        hdr_fmt = wb.add_format({"bold": True})
    else:
        sio = io.StringIO()
        writer = csv.writer(sio)
        writer.writerow(headers)

    for start in range(0, total, BATCH):
        chunk = ids[start:start + BATCH]
        rows = call(a, MODEL, "export_data", [chunk, paths])["datas"]
        rows = [[flat(v) for v in row] for row in rows]

        if is_xlsx:
            if start == 0:
                # أول دفعة: نظبط العرض والتجميد والهيدر قبل كتابة الصفوف
                sample = rows[:2000]
                for i, h in enumerate(headers):
                    longest = max((len(str(r[i])) for r in sample), default=0)
                    ws.set_column(i, i, min(max(len(h), longest) + 2, 50))
                ws.freeze_panes(1, 0)
                ws.write_row(0, 0, headers, hdr_fmt)
            for row in rows:
                if n >= EXCEL_MAX:
                    truncated = True
                    break
                n += 1
                for c, v in enumerate(row):
                    if v is None or v == "":
                        continue
                    if isinstance(v, str):
                        ws.write_string(n, c, v)
                    else:
                        ws.write(n, c, v)
        else:
            writer.writerows(rows)
            n += len(rows)

        if len(preview) < 200:
            preview += rows[:200 - len(preview)]

        done = min(start + len(chunk), total)
        progress.progress(done / total, text=f"تم سحب وكتابة {done:,} من {total:,}")
        if truncated:
            break

    progress.progress(1.0, text="بنقفل الملف...")
    if is_xlsx:
        wb.close()
        data = buf.getvalue()
    else:
        data = sio.getvalue().encode("utf-8-sig")

    return dict(data=data, rows=n, preview=pd.DataFrame(preview, columns=headers),
                truncated=truncated)


def unwrap_xlsx(file_bytes):
    out, count = io.BytesIO(), 0
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zin, \
         zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED, compresslevel=1, allowZip64=True) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/styles.xml":
                root = etree.fromstring(data)
                for al in root.iter(f"{{{NS}}}alignment"):
                    if al.get("wrapText") in ("1", "true"):
                        al.attrib.pop("wrapText")
                        count += 1
                data = etree.tostring(root, encoding="UTF-8", xml_declaration=True)
            zout.writestr(item, data)
    return out.getvalue(), count


# ======================= Callbacks =======================
def logout():
    for k in list(st.session_state.keys()):
        del st.session_state[k]


def apply_template():
    name = st.session_state.get("tpl")
    tpls = st.session_state["data"]["templates"]
    if name in tpls:
        st.session_state["cols"] = list(tpls[name])
        st.session_state.pop("result", None)


# ======================= الشاشات =======================
def login_screen():
    st.markdown("## 👋 أهلًا بيك")
    st.caption("سجّل دخولك بحسابك في أودو عشان تسحب بياناتك.")
    with st.form("login"):
        user = st.text_input("اسم المستخدم")
        pwd = st.text_input("كلمة المرور", type="password")
        ok = st.form_submit_button("تسجيل الدخول", type="primary", use_container_width=True)
    if ok:
        if not user or not pwd:
            st.warning("اكتب اسم المستخدم وكلمة المرور.")
            return
        try:
            with st.spinner("بنتأكد من البيانات..."):
                uid = rpc("common", "login", DB, user, pwd)
            if not uid:
                st.error("اسم المستخدم أو كلمة المرور غلط.")
                return
            st.session_state["auth"] = dict(user=user, pwd=pwd, uid=uid)
            st.rerun()
        except Exception as e:
            st.error(f"مقدرناش نتصل بأودو: {e}")


def main_screen():
    a = st.session_state["auth"]

    with st.sidebar:
        st.markdown(f"👤 **{a['user']}**")
        st.button("تسجيل خروج", on_click=logout)

    if "data" not in st.session_state:
        try:
            with st.spinner("لحظات... بنجهّز البيانات"):
                st.session_state["data"] = load_everything(a)
        except Exception as e:
            st.error(f"حصل خطأ أثناء التحميل: {e}")
            st.button("🔄 جرّب تاني", on_click=lambda: st.session_state.pop("data", None))
            return
    d = st.session_state["data"]
    groups, display, templates = d["groups"], d["display"], d["templates"]

    st.markdown("## 📥 سحب بيانات Claims")

    # ---------------- الخطوة 1 ----------------
    with st.container(border=True):
        st.markdown("### 1️⃣ اختار الشركة")
        sel = st.multiselect(
            "الشركات", list(range(len(groups))),
            format_func=lambda i: f"{groups[i]['name']}  ({groups[i]['count']:,})",
            placeholder="اختار شركة أو أكتر...",
            label_visibility="collapsed",
        )
        total = sum(groups[i]["count"] for i in sel)
        if sel:
            st.caption(f"عدد السجلات: **{total:,}**")

    # ---------------- الخطوة 2 ----------------
    with st.container(border=True):
        st.markdown("### 2️⃣ اختار الأعمدة")
        if templates:
            st.selectbox("📋 قالب جاهز (اختياري)", ["— من غير قالب —"] + list(templates.keys()),
                         key="tpl", on_change=apply_template)

        q = st.text_input("🔎 دوّر على عمود وضيفه", key="colq",
                          placeholder="اكتب جزء من اسم العمود...").strip().lower()
        selected = list(st.session_state.get("cols", []))
        matches = [n for n in d["field_names"] if q in display[n].lower()][:100]
        options = selected + [n for n in matches if n not in selected]

        cols = st.multiselect(
            "الأعمدة المختارة", options, key="cols",
            format_func=lambda p: display.get(p) or path_label(p, d["labels"]),
            placeholder="اختار الأعمدة...",
        )
        st.caption("بيظهر أول 100 عمود، استخدم البحث عشان توصل لأي عمود تاني.")

    # ---------------- الخطوة 3 ----------------
    with st.container(border=True):
        st.markdown("### 3️⃣ اسحب ونزّل")
        c1, c2 = st.columns(2)
        fmt = c1.radio("شكل الملف", ["Excel", "CSV"], horizontal=True)
        test = c2.checkbox("تجربة سريعة (أول 5000 سجل بس)")

        if fmt == "Excel" and total > EXCEL_MAX and not test:
            st.warning("العدد أكبر من حد الإكسيل (مليون صف)، هننزّله CSV.")
            fmt = "CSV"

        go = st.button("🚀 ابدأ السحب", type="primary", use_container_width=True,
                       disabled=not (sel and cols))
        if not (sel and cols):
            st.caption("اختار شركة وعمود واحد على الأقل عشان الزرار يشتغل.")

        if go:
            st.session_state.pop("result", None)
            try:
                keys = [groups[i]["key"] for i in sel]
                domain = BASE_DOMAIN + [[d["portfolio"], "in", keys]]
                headers, seen = [], {}
                for p in cols:
                    h = path_label(p, d["labels"])
                    seen[h] = seen.get(h, 0) + 1
                    headers.append(h if seen[h] == 1 else f"{h}.{seen[h]}")

                bar = st.progress(0.0, text="جاري السحب...")
                res = build_file(a, domain, cols, headers, TEST_LIMIT if test else None, fmt, bar)
                bar.empty()

                if not res:
                    st.warning("مفيش بيانات مطابقة.")
                else:
                    name = groups[sel[0]]["name"] if len(sel) == 1 else "claims"
                    name = re.sub(r'[\\/:*?"<>|]+', "", name).strip() or "claims"
                    ext = "xlsx" if fmt == "Excel" else "csv"
                    res.update(
                        file=f"{name}_{datetime.now():%Y-%m-%d_%H%M}.{ext}",
                        mime=("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                              if ext == "xlsx" else "text/csv"),
                    )
                    st.session_state["result"] = res
            except Exception as e:
                st.error(f"حصل خطأ: {e}")

    # ---------------- النتيجة ----------------
    r = st.session_state.get("result")
    if r:
        st.success(f"✅ تم سحب {r['rows']:,} سجل، الملف جاهز.")
        if r["truncated"]:
            st.warning("وصلنا لحد صفوف الإكسيل، الباقي اتقطع. استخدم CSV لو عايز كل السجلات.")
        st.download_button("⬇️ نزّل الملف", r["data"], file_name=r["file"], mime=r["mime"],
                           type="primary", use_container_width=True)
        with st.expander("👀 معاينة أول 200 صف"):
            st.dataframe(r["preview"], use_container_width=True)


def unwrap_screen():
    st.markdown("## 🧹 فك Wrap من ملف Excel جاهز")
    st.caption("ارفع أي ملف Excel وهيرجعلك نسخة من غير Wrap Text.")
    up = st.file_uploader("ارفع ملف xlsx", type="xlsx")
    if up:
        fixed, n = unwrap_xlsx(up.read())
        st.success(f"تم إلغاء Wrap Text من {n} تنسيق.")
        st.download_button("⬇️ نزّل الملف بعد التعديل", fixed, file_name=f"nowrap_{up.name}",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary")


# ======================= التشغيل =======================
tab1, tab2 = st.tabs(["📥 سحب البيانات", "🧹 فك Wrap من ملف"])
with tab1:
    if "auth" in st.session_state:
        main_screen()
    else:
        login_screen()
with tab2:
    unwrap_screen()