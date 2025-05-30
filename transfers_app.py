import re
import streamlit as st
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import zipfile
from io import BytesIO
import os
import socket
import time
import threading
import math
import webbrowser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Check if running on Windows (for local development)
WINDOWS_ENV = False
try:
    import win32com.client
    import pythoncom
    WINDOWS_ENV = True
except ImportError:
    pass

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', name)

def safe_str(text):
    return str(text).replace("–", "-").encode('latin-1', errors='replace').decode('latin-1')

def send_email(recipient, subject, body, attachment_path=None):
    """Generic email function that works both locally (Outlook) and in cloud (SMTP)"""
    
    # Cloud environment - use SMTP
    if not WINDOWS_ENV:
        try:
            # Configure these in your Streamlit secrets
            smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
            smtp_port = st.secrets.get("SMTP_PORT", 587)
            smtp_username = st.secrets.get("SMTP_USERNAME")
            smtp_password = st.secrets.get("SMTP_PASSWORD")
            
            if not all([smtp_server, smtp_username, smtp_password]):
                st.warning("Email configuration incomplete. Please set SMTP credentials in secrets.")
                return False

            msg = MIMEMultipart()
            msg['From'] = smtp_username
            msg['To'] = recipient if isinstance(recipient, str) else ", ".join(recipient)
            msg['Subject'] = subject
            msg.attach(MIMEText(body, 'plain'))

            if attachment_path and os.path.exists(attachment_path):
                with open(attachment_path, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={os.path.basename(attachment_path)}',
                )
                msg.attach(part)

            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_username, smtp_password)
                server.send_message(msg)
            
            return True
            
        except Exception as e:
            st.error(f"Failed to send email via SMTP: {str(e)}")
            return False
    
    # Local environment - use Outlook
    else:
        try:
            pythoncom.CoInitialize()
            if attachment_path and not os.path.exists(attachment_path):
                st.error(f"❌ Attachment not found: {attachment_path}")
                return False

            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            mail.To = "; ".join(recipient) if isinstance(recipient, list) else recipient
            mail.Subject = subject
            mail.Body = body
            if attachment_path:
                mail.Attachments.Add(os.path.abspath(attachment_path))
            mail.Display()  # Opens the email draft without sending
            return True
        except Exception as e:
            st.exception(e)
            return False

def is_port_open(host, port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex((host, port)) == 0

def open_browser_once():
    for _ in range(30):
        if is_port_open("localhost", 8501):
            if not os.environ.get("BROWSER_OPENED"):
                webbrowser.open("http://localhost:8501")
                os.environ["BROWSER_OPENED"] = "1"
            break
        time.sleep(1)

if __name__ == "__main__":
    threading.Thread(target=open_browser_once).start()
    os.environ["BROWSER"] = "none"

def save_excel_report(df, filename):
    try:
        os.makedirs("reports", exist_ok=True)
        file_path = f"reports/{sanitize_filename(safe_str(filename))}"
        df.to_excel(file_path, index=False)
        return file_path
    except Exception as e:
        st.error(f"Failed to generate Excel: {str(e)}")
        return None

def generate_pdf_report(merchant_id, merchant_name, iban, trx_df, total_payment, date_from, date_to, output_file):
    try:
        trx_df = trx_df.copy()
        
        col_mapping = {
            'terminalid': 'terminal_id',
            'terminal id': 'terminal_id',
            'cardnumber': 'card_number',
            'card number': 'card_number',
            'maskedcardnumber': 'card_number',
            'masked card number': 'card_number',
            'authcode': 'auth_code',
            'auth code': 'auth_code',
            'trxcurrency': 'transaction_currency',
            'transactioncurrency': 'transaction_currency',
            'transaction currency': 'transaction_currency'
        }

        trx_df.columns = [col_mapping.get(col.lower().strip(), col) for col in trx_df.columns]

        if 'terminal_id' not in trx_df.columns:
            trx_df['terminal_id'] = 'All Terminals'
        if 'card_number' not in trx_df.columns:
            trx_df['card_number'] = 'N/A'
        if 'transaction_currency' not in trx_df.columns:
            trx_df['transaction_currency'] = 'IQD'
        if 'auth_code' not in trx_df.columns:
            trx_df['auth_code'] = 'N/A'
        if 'process_date' not in trx_df.columns:
            trx_df['process_date'] = datetime.now()
        if 'vtr_dcc_amou' not in trx_df.columns:
            trx_df['vtr_dcc_amou'] = 0
        if 'merchant_commission' not in trx_df.columns:
            trx_df['merchant_commission'] = 0

        class PDF(FPDF):
            def header(self):
                self.set_font("Arial", 'B', 12)
                self.cell(0, 8, safe_str(merchant_name.upper()), ln=True)
                self.set_font("Arial", '', 10)
                self.cell(0, 6, f"Merchant ID: {merchant_id}", ln=True)
                self.cell(0, 6, f"Account No: {iban}", ln=True)
                self.cell(0, 6, f"From: {date_from}  To: {date_to}", ln=True)
                self.cell(0, 6, f"Report Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True)
                self.ln(4)

            def footer(self):
                self.set_y(-15)
                self.set_font("Arial", 'I', 8)
                self.cell(0, 10, f"Page {self.page_no()}", 0, 0, 'C')

        pdf = PDF(orientation='P', unit='mm', format='A4')
        pdf.set_auto_page_break(auto=True, margin=10)
        pdf.add_page()
        pdf.set_font("Arial", size=8)

        max_width = 190
        headers = ["Date", "Card", "Type", "Curr", "Amount", "Comm", "Net", "MDR", "Auth"]
        widths = [18, 30, 12, 8, 18, 18, 18, 15, 20]
        grand_total_dcc, grand_total_comm, grand_total_net = 0, 0, 0

        grouped = trx_df.groupby('terminal_id')

        for terminal_id, group in grouped:
            pdf.set_font("Arial", 'B', 9)
            terminal_str = str(int(float(terminal_id))) if str(terminal_id).replace(".", "", 1).isdigit() else str(terminal_id)
            pdf.cell(0, 6, f"Terminal ID: {terminal_str}", ln=True)
            pdf.ln(2)

            x_offset = (190 - sum(widths)) / 2
            pdf.set_x(x_offset)
            for h, w in zip(headers, widths):
                pdf.cell(w, 6, h, border=1, align='C')
            pdf.ln()

            total_dcc, total_comm, total_net = 0, 0, 0

            for _, row in group.iterrows():
                card = str(row.get('card_number', 'N/A')).strip()
                curr = 'IQD' if 'dinar' in str(row.get('transaction_currency', 'IQD')).lower() else str(row.get('transaction_currency', 'IQD')).upper()[:3]
                auth = str(row.get('auth_code', 'N/A')).strip()
                date = row['process_date'].strftime('%d-%m-%Y')
                dcc = float(row.get('vtr_dcc_amou', 0))
                comm = float(row.get('merchant_commission', 0))
                net = dcc + comm
                mdr = f"{abs(comm)/dcc*100:.1f}%" if dcc != 0 else "0.0%"

                row_data = [
                    date,
                    card,
                    "SALE",
                    curr,
                    f"{dcc:,.2f}",
                    f"{comm:,.2f}",
                    f"{net:,.2f}",
                    mdr,
                    auth
                ]

                aligns = ['L', 'L', 'C', 'C', 'R', 'R', 'R', 'C', 'C']
                pdf.set_x(x_offset)
                pdf.set_font("Arial", size=7)
                for val, w, a in zip(row_data, widths, aligns):
                    pdf.cell(w, 5, str(val), border=1, align=a)
                pdf.ln()

                total_dcc += dcc
                total_comm += comm
                total_net += net

            pdf.set_font("Arial", 'B', 8)
            pdf.set_x(x_offset)
            pdf.cell(sum(widths[:4]), 5, "Subtotal", border=1, align='R')
            pdf.cell(widths[4], 5, f"{total_dcc:,.2f}", border=1, align='R')
            pdf.cell(widths[5], 5, f"{total_comm:,.2f}", border=1, align='R')
            pdf.cell(widths[6], 5, f"{total_net:,.2f}", border=1, align='R')
            pdf.cell(widths[7] + widths[8], 5, "", border=1)
            pdf.ln(8)

            grand_total_dcc += total_dcc
            grand_total_comm += total_comm
            grand_total_net += total_net

        pdf.set_font("Arial", 'B', 9)
        pdf.cell(0, 6, "TOTAL SUMMARY", ln=True)
        pdf.cell(40, 6, "Total Amount:", ln=False)
        pdf.cell(0, 6, f"{grand_total_dcc:,.2f} IQD", ln=True)
        pdf.cell(40, 6, "Total Commission:", ln=False)
        pdf.cell(0, 6, f"{grand_total_comm:,.2f} IQD", ln=True)
        pdf.cell(40, 6, "Total Net Amount:", ln=False)
        pdf.cell(0, 6, f"{grand_total_net:,.2f} IQD", ln=True)

        os.makedirs("reports", exist_ok=True)
        file_path = f"reports/{sanitize_filename(safe_str(output_file))}"
        pdf.output(file_path)
        return file_path

    except Exception as e:
        st.error(f"Failed to generate PDF: {str(e)}")
        return None

# --- Main Streamlit App ---
st.set_page_config(page_title="Transfers App", layout="wide")
st.title("💳💸 Transfers App – Version 1.0")

# File uploaders
payments_file = st.file_uploader("📥 Upload Payments Excel", type=["xlsx"])
transactions_file = st.file_uploader("📥 Upload Transactions File", type=["csv", "xlsx"])
filter_file = st.file_uploader("📤 Optional: Upload Merchant Filter Excel", type=["xlsx"])
global_to_date = st.date_input("📅 Global 'To' Date (optional auto-match trigger)", value=None)

@st.cache_data
def load_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.astype(str).str.strip().str.lower()
    return df

@st.cache_data
def load_transactions(file):
    try:
        if file.name.lower().endswith('.csv'):
            df = pd.read_csv(file, encoding="ISO-8859-1")
        else:
            df = pd.read_excel(file)

        col_map = {
            'merchant id': 'merchant_id',
            'process date': 'process_date',
            'merchant commission': 'merchant_commission',
            'vtr dcc amou': 'vtr_dcc_amou',
            'transaction currency': 'transaction_currency',
            'card number': 'card_number',
            'masked card number': 'card_number',
            'auth code': 'auth_code',
            'terminal id': 'terminal_id'
        }

        df.columns = df.columns.str.lower().str.strip()
        df.rename(columns={k.lower(): v for k, v in col_map.items()}, inplace=True)

        for col in ['merchant_id', 'process_date', 'merchant_commission', 'vtr_dcc_amou', 'card_number', 'auth_code']:
            if col not in df.columns:
                df[col] = None

        df["merchant_id"] = df["merchant_id"].astype(str)
        df["process_date"] = pd.to_datetime(df["process_date"], errors="coerce")
        df["vtr_dcc_amou"] = pd.to_numeric(df["vtr_dcc_amou"], errors="coerce").fillna(0)
        df["merchant_commission"] = pd.to_numeric(df["merchant_commission"], errors="coerce").fillna(0)

        return df[df["process_date"] >= pd.to_datetime("today") - pd.Timedelta(days=90)]
    except Exception as e:
        st.error(f"❌ Failed to load transaction file: {e}")
        return pd.DataFrame()

if payments_file and transactions_file:
    payments_df = load_excel(payments_file)
    transactions_df = load_transactions(transactions_file)

    payments_df.columns = payments_df.columns.str.strip().str.lower()
    payments_df.rename(columns={
        'merchant name': 'merchant_name',
        'merchant id': 'merchant_id',
        'payment': 'payment_amount',
        'amount': 'payment_amount',
        'iban': 'iban'
    }, inplace=True)

    payments_df["merchant_id"] = payments_df["merchant_id"].astype(str)
    payments_df["merchant_name"] = payments_df.get('merchant_name', '')

    def get_name_from_trx(mid):
        name = transactions_df.loc[transactions_df["merchant_id"] == mid, "merchant_name"]
        return name.iloc[0] if not name.empty else f"Merchant {mid}"

    payments_df["merchant_name"] = payments_df.apply(
        lambda row: row["merchant_name"] if pd.notna(row["merchant_name"]) and str(row["merchant_name"]).strip() != ''
        else get_name_from_trx(row["merchant_id"]),
        axis=1
    )

    if filter_file:
        filter_df = load_excel(filter_file)
        filter_df.columns = filter_df.columns.str.strip().str.lower()
        filter_df.rename(columns={"merchant id": "merchant_id"}, inplace=True)
        filter_df["merchant_id"] = filter_df["merchant_id"].astype(str)
        payments_df = payments_df[payments_df["merchant_id"].isin(filter_df["merchant_id"])]
        transactions_df = transactions_df[transactions_df["merchant_id"].isin(filter_df["merchant_id"])]
        merchant_emails = dict(zip(filter_df["merchant_id"], filter_df["email"])) if "email" in filter_df.columns else {}
    else:
        merchant_emails = {}

    if "merchant_dates" not in st.session_state:
        st.session_state.merchant_dates = {}

    if st.button("🔄"):
        st.session_state.calculate_trigger = True

    all_pdf_paths = []
    for _, row in payments_df.iterrows():
        with st.container():
            merchant_id = str(row["merchant_id"])
            merchant_name = str(row["merchant_name"])
            iban = row["iban"]
            payment_amount = float(row["payment_amount"])

            from_key = f"from_{merchant_id}"
            to_key = f"to_{merchant_id}"
            default_from = st.session_state.merchant_dates.get(merchant_id, {}).get("from", datetime.today())
            default_to = st.session_state.merchant_dates.get(merchant_id, {}).get("to", datetime.today())

            cols = st.columns([2, 2, 2, 2, 2, 5])
            cols[0].markdown(f"**{merchant_name}**")
            cols[1].markdown(f"**{merchant_id}**")
            from_date = cols[2].date_input("From", value=default_from, key=from_key)
            to_date = cols[3].date_input("To", value=default_to, key=to_key)
            st.session_state.merchant_dates[merchant_id] = {"from": from_date, "to": to_date}

            user_modified_dates = (
                st.session_state.merchant_dates.get(merchant_id, {}).get("from") != default_from or
                st.session_state.merchant_dates.get(merchant_id, {}).get("to") != default_to
            )

            if global_to_date and not user_modified_dates:
                merchant_trx = transactions_df[transactions_df["merchant_id"] == merchant_id].copy()
                merchant_trx = merchant_trx.dropna(subset=["vtr_dcc_amou", "merchant_commission"])
                merchant_trx["net_amount"] = merchant_trx["vtr_dcc_amou"] + merchant_trx["merchant_commission"]
                merchant_trx = merchant_trx.sort_values("process_date")

                trx = pd.DataFrame()
                valid_dates = merchant_trx[merchant_trx["process_date"] <= pd.to_datetime(global_to_date)]["process_date"].dropna().sort_values().unique()

                for start in valid_dates:
                    start_ts = pd.to_datetime(start)
                    temp_trx = merchant_trx[
                        (merchant_trx["process_date"] >= start_ts) &
                        (merchant_trx["process_date"] <= pd.to_datetime(global_to_date))
                    ]
                    if math.isclose(temp_trx["net_amount"].sum(), payment_amount, rel_tol=1e-5, abs_tol=0.05):
                        trx = temp_trx
                        from_date = trx["process_date"].min().date()
                        to_date = trx["process_date"].max().date()
                        st.session_state.merchant_dates[merchant_id] = {"from": from_date, "to": to_date}
                        break
            else:
                trx = transactions_df[
                    (transactions_df["merchant_id"] == merchant_id) &
                    (transactions_df["process_date"] >= pd.to_datetime(from_date)) &
                    (transactions_df["process_date"] <= pd.to_datetime(to_date))
                ].copy()

            if {"vtr_dcc_amou", "merchant_commission"}.issubset(trx.columns):
                trx = trx.dropna(subset=["vtr_dcc_amou", "merchant_commission"])
                trx["net_amount"] = trx["vtr_dcc_amou"] + trx["merchant_commission"]
            else:
                trx = pd.DataFrame(columns=["vtr_dcc_amou", "merchant_commission", "net_amount"])
                trx["net_amount"] = 0

            net_total = trx["net_amount"].sum()
            cols[4].markdown(f"💵 `{payment_amount:,.2f}`<br>🧾 `{net_total:,.2f}`", unsafe_allow_html=True)

            if math.isclose(net_total, payment_amount, rel_tol=1e-5, abs_tol=0.05):
                safe_merchant_name = sanitize_filename(merchant_name.replace(" ", "_"))
                file_name = f"{merchant_id}_{safe_merchant_name}.pdf"

                try:
                    file_path = generate_pdf_report(
                        merchant_id,
                        merchant_name,
                        iban,
                        trx,
                        net_total,
                        from_date.strftime("%d/%m/%Y"),
                        to_date.strftime("%d/%m/%Y"),
                        file_name
                    )

                    if file_path and os.path.exists(file_path):
                        all_pdf_paths.append(file_path)
                        excel_file_name = f"{merchant_id}_{safe_merchant_name}.xlsx"
                        excel_file_path = save_excel_report(trx, excel_file_name)

                        # Create button container in the last column
                        with cols[5]:
                            # Create 5 columns for the buttons
                            btn_col1, btn_col2, btn_col3, btn_col4, btn_col5 = st.columns(5)
                            
                            # PDF Download
                            with btn_col1:
                                with open(file_path, "rb") as f:
                                    st.download_button(
                                        label="📄 PDF",
                                        data=f,
                                        file_name=file_name,
                                        mime="application/pdf",
                                        key=f"dl_pdf_{merchant_id}"
                                    )
                            
                            # Excel Download
                            with btn_col2:
                                with open(excel_file_path, "rb") as ef:
                                    st.download_button(
                                        label="📊 Excel",
                                        data=ef,
                                        file_name=excel_file_name,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key=f"dl_excel_{merchant_id}"
                                    )
                            
                            # Send PDF
                            with btn_col3:
                                if merchant_id in merchant_emails:
                                    email = merchant_emails[merchant_id]
                                    email_subject = f"Settlement Report for {merchant_name}"
                                    email_body = (
                                        f"Dear,\n\n"
                                        f"Please find attached your transaction settlement report.\n\n"
                                        f"Merchant ID: {merchant_id}\n"
                                        f"Date Range: {from_date.strftime('%d/%m/%Y')} to {to_date.strftime('%d/%m/%Y')}\n"
                                        f"Settled Amount: {net_total:,.2f} IQD\n\n"
                                        "Best regards"
                                    )
                                    if st.button("📧 PDF", key=f"email_pdf_{merchant_id}"):
                                        if send_email(email, email_subject, email_body, file_path):
                                            st.success(f"PDF sent to {email}")
                                        else:
                                            st.error(f"Failed to send PDF")
                            
                            # Send Excel
                            with btn_col4:
                                if merchant_id in merchant_emails:
                                    excel_subject = f"Settlement Excel Report for {merchant_name}"
                                    excel_body = (
                                        f"Dear,\n\n"
                                        f"Please find attached your Excel settlement report.\n\n"
                                        f"Merchant ID: {merchant_id}\n"
                                        f"Date Range: {from_date.strftime('%d/%m/%Y')} to {to_date.strftime('%d/%m/%Y')}\n"
                                        f"Settled Amount: {net_total:,.2f} IQD\n\n"
                                        "Best regards"
                                    )
                                    if st.button("📧 Excel", key=f"email_excel_{merchant_id}"):
                                        if send_email(email, excel_subject, excel_body, excel_file_path):
                                            st.success(f"Excel sent to {email}")
                                        else:
                                            st.error(f"Failed to send Excel")
                            
                            # Send Both
                            with btn_col5:
                                if merchant_id in merchant_emails:
                                    if st.button("📧 Both", key=f"email_both_{merchant_id}"):
                                        if WINDOWS_ENV:
                                            try:
                                                pythoncom.CoInitialize()
                                                outlook = win32com.client.Dispatch("Outlook.Application")
                                                mail = outlook.CreateItem(0)
                                                mail.To = email
                                                mail.Subject = f"Complete Settlement Report for {merchant_name}"
                                                mail.Body = (
                                                    f"Dear,\n\n"
                                                    f"Please find attached both PDF and Excel versions of your settlement report.\n\n"
                                                    f"Merchant ID: {merchant_id}\n"
                                                    f"Date Range: {from_date.strftime('%d/%m/%Y')} to {to_date.strftime('%d/%m/%Y')}\n"
                                                    f"Settled Amount: {net_total:,.2f} IQD\n\n"
                                                    "Best regards"
                                                )
                                                mail.Attachments.Add(os.path.abspath(file_path))
                                                mail.Attachments.Add(os.path.abspath(excel_file_path))
                                                mail.Display()
                                                st.success(f"Email draft created")
                                            except Exception as e:
                                                st.error(f"Failed to create email: {str(e)}")
                                        else:
                                            # Cloud version - send both attachments via SMTP
                                            combined_subject = f"Complete Settlement Report for {merchant_name}"
                                            combined_body = (
                                                f"Dear,\n\n"
                                                f"Please find attached both PDF and Excel versions of your settlement report.\n\n"
                                                f"Merchant ID: {merchant_id}\n"
                                                f"Date Range: {from_date.strftime('%d/%m/%Y')} to {to_date.strftime('%d/%m/%Y')}\n"
                                                f"Settled Amount: {net_total:,.2f} IQD\n\n"
                                                "Best regards"
                                            )
                                            if send_email(email, combined_subject, combined_body, file_path) and \
                                               send_email(email, combined_subject, combined_body, excel_file_path):
                                                st.success(f"Both files sent to {email}")
                                            else:
                                                st.error(f"Failed to send both files")

                    else:
                        st.error(f"❌ Failed to generate PDF for {merchant_name}")
                except Exception as e:
                    st.error(f"❌ PDF generation failed: {str(e)}")
            else:
                cols[5].error("❌ Mismatch")

    if all_pdf_paths:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for path in all_pdf_paths:
                if os.path.exists(path):
                    zip_file.write(path, os.path.basename(path))

        st.download_button(
            label="📦 Download All PDFs (ZIP)",
            data=zip_buffer.getvalue(),
            file_name="all_matched_reports.zip",
            mime="application/zip"
        )