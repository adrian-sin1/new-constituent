import streamlit as st
import pandas as pd
import re
import csv
import io
import os
import xlsxwriter
import tkinter as tk
from selenium.common.exceptions import NoSuchElementException
from streamlit.runtime.scriptrunner import RerunException


# ------------------ EMAIL PARSER FUNCTION ------------------

def extract_replies_with_senders(body, csv_email):
    pattern = re.compile(
        r'(?=^From:|^On .+? wrote:|^-----Original Message-----)',
        re.IGNORECASE | re.MULTILINE
    )
    chunks = pattern.split(body.strip())
    results = []
    last_sender = csv_email

    for i, chunk in enumerate(chunks):
        chunk = chunk.strip()
        if not chunk or len(chunk.splitlines()) < 2:
            continue

        sender = None
        if i == 0:
            sender = csv_email
        else:
            match_from = re.search(r'^From:\s*(.*)', chunk, re.IGNORECASE | re.MULTILINE)
            if match_from:
                sender = match_from.group(1).strip()
            else:
                match_wrote = re.search(r'On .+? (.+?) <(.+?)> wrote:', chunk, re.IGNORECASE)
                if match_wrote:
                    name = match_wrote.group(1).strip()
                    email_addr = match_wrote.group(2).strip()
                    sender = f"{name} <{email_addr}>"

        sender = sender or last_sender
        last_sender = sender
        results.append((sender, chunk))

    return results

# ------------------ CREDENTIAL ENTRY (TKINTER) ------------------

def get_credentials_from_tkinter():
    credentials = {}

    def submit():
        credentials["username"] = username_var.get()
        credentials["password"] = password_var.get()
        credentials["auto_click"] = auto_click_var.get()
        root.destroy()

    root = tk.Tk()
    root.title("Council Connect Login")
    root.geometry("350x200")
    root.attributes('-topmost', True)

    tk.Label(root, text="Council ID:").pack(pady=(10, 0))
    username_var = tk.StringVar()
    tk.Entry(root, textvariable=username_var, width=40).pack()

    tk.Label(root, text="Password:").pack(pady=(10, 0))
    password_var = tk.StringVar()
    tk.Entry(root, textvariable=password_var, show="*", width=40).pack()

    auto_click_var = tk.BooleanVar()
    tk.Checkbutton(root, text="Auto-click 'Save'", variable=auto_click_var).pack(pady=10)

    tk.Button(root, text="Start", command=submit, width=15).pack(pady=10)
    root.mainloop()
    return credentials if credentials else None

# ------------------ STREAMLIT SETUP ------------------

st.set_page_config(page_title="📬 New Constituent Emails", layout="wide")
st.title("📬 New Constituent Emails")
st.markdown("Upload a raw exported email CSV file. This app will parse and group replies by thread (subject). Select threads to export or upload to Council Connect.")

uploaded_file = st.file_uploader("📥 Upload raw exported email CSV", type="csv")

if not uploaded_file:
    st.warning("Please upload a raw email export file.")
    st.stop()

# ------------------ PARSE ------------------

reader = csv.DictReader(io.StringIO(uploaded_file.getvalue().decode("ISO-8859-1")))
rows = []

for row in reader:
    name = row.get("To: (Name)", "").strip(" '\"")
    email = row.get("To: (Address)", "").strip(" '\"")
    subject = row.get("Subject", "").strip(" '\"")
    body = row.get("Body", "")

    if email.lower().startswith("/o=nycc/ou=exchange"):
        match_emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b', body)
        email = match_emails[0] if match_emails else "/o=NYCC/ou=Exchange Administrative"

    replies = extract_replies_with_senders(body, email)

    for sender, reply_text in replies:
        rows.append({
            "Name": name,
            "Email": email,
            "Subject": subject,
            
            "Reply": reply_text
        })

df = pd.DataFrame(rows)
df["__order"] = range(len(df))
df["Subject"] = df.get("Subject", "No Subject").fillna("No Subject")
grouped = df.groupby("Subject", sort=False)

st.success(f"✅ Parsed {len(df)} replies from {len(grouped)} threads.")

# ------------------ THREAD SELECTION ------------------

total_groups = len(grouped)
for i in range(1, total_groups + 1):
    key = f"select_{i}"
    if key not in st.session_state:
        st.session_state[key] = False

selected_groups = []
for i, (subject, group) in enumerate(grouped, start=1):
    group = group.sort_values("__order")
    first_name = group.iloc[0].get("Name", "Unknown")
    if st.session_state.get(f"select_{i}", False):
        selected_groups.append(group)

result_df = pd.concat(selected_groups, ignore_index=True).drop(columns=["__order"], errors="ignore") if selected_groups else None

# ------------------ DOWNLOAD OPTIONS ------------------

st.subheader("📤 Download Options")
export_format = st.radio("Choose export format:", ["CSV", "Excel (.xlsx)", "Notepad (.txt)"], horizontal=True)
download_placeholder = st.empty()

if result_df is not None:
    if export_format == "CSV":
        csv_data = result_df.to_csv(index=False, quoting=csv.QUOTE_ALL)
        download_placeholder.download_button("⬇️ Download CSV", data=csv_data,
                                             file_name="filtered_output.csv", mime="text/csv")
    elif export_format == "Excel (.xlsx)":
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False, sheet_name='Emails')
        excel_buffer.seek(0)
        download_placeholder.download_button("⬇️ Download Excel File", data=excel_buffer.getvalue(),
                                             file_name="filtered_output.xlsx",
                                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif export_format == "Notepad (.txt)":
        txt_output = ""
        for _, row in result_df.iterrows():
            txt_output += f"Name: {row.get('Name', '')}\n"
            txt_output += f"Email: {row.get('Email', '')}\n"
            
            txt_output += f"Subject: {row.get('Subject', '')}\n"
            txt_output += f"Reply:\n{row.get('Reply', '').strip()}\n"
            txt_output += "-" * 40 + "\n"
        download_placeholder.download_button("⬇️ Download Notepad File", data=txt_output,
                                             file_name="filtered_output.txt", mime="text/plain")
else:
    download_placeholder.empty()

# ------------------ UPLOAD TO COUNCIL CONNECT ------------------

st.subheader("🏢 Upload to Council Connect")
launch_login = st.button("🔐 Login Credentials")

if launch_login:
    creds = get_credentials_from_tkinter()

    if not creds:
        st.error("Login cancelled.")
    elif result_df is None or result_df.empty:
        st.error("Please select at least one thread before uploading.")
    else:
        with st.spinner("Launching browser and submitting entries..."):
            from upload import upload_to_council_connect
            DRIVER_PATH = os.path.join(os.getcwd(), "msedgedriver.exe")

            try:
                upload_to_council_connect(result_df, creds["username"], creds["password"], creds["auto_click"], DRIVER_PATH)
                st.success("✅ All selected threads uploaded to Council Connect.")
                
                st.rerun()

            except Exception as e:
                st.error(f"❌ Upload failed: {str(e)}")

# ------------------ SEARCH + THREAD VIEW ------------------

st.subheader("🔍 Search Threads")
search_query = st.text_input("Search by subject or name:", "").strip().lower()

st.subheader("📑 Email Threads")

select_all_state = st.checkbox("🔘 Select All Threads", key="select_all_toggle")
if "prev_select_all" not in st.session_state:
    st.session_state.prev_select_all = False
if select_all_state != st.session_state.prev_select_all:
    for i in range(1, total_groups + 1):
        st.session_state[f"select_{i}"] = select_all_state
    st.session_state.prev_select_all = select_all_state

matches_found = False

for i, (subject, group) in enumerate(grouped, start=1):
    group = group.sort_values("__order")
    first_name = group.iloc[0].get("Name", "Unknown")
    searchable_text = f"{first_name} {subject}".lower()

    if search_query and search_query not in searchable_text:
        continue

    matches_found = True
    expander_title = f"{i}. {first_name} | {subject} ({len(group)} replies)"
    col1, col2 = st.columns([0.05, 0.95])

    with col1:
        st.checkbox("✔", key=f"select_{i}", label_visibility="collapsed")
    with col2:
        with st.expander(expander_title, expanded=False):
            for _, row in group.iterrows():
                st.markdown(f"""
**Name**: {row.get("Name", "")}  
**Email**: {row.get("Email", "")}  
**Reply:**  
```text
{row.get("Reply", "").strip()}
```""", unsafe_allow_html=True)

if not matches_found:
    st.info("No threads matched your search.")