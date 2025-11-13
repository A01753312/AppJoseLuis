import streamlit as st
import pandas as pd
import base64
import requests

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

st.set_page_config(page_title="Correo masivo OAuth", page_icon="📧")

st.title("📧 Envío automático de correos (Google)")

# ================== CONFIG GOOGLE ==================
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
GOOGLE_REDIRECT_URI = st.secrets["GOOGLE_REDIRECT_URI"]

google_client_config = {
    "web": {
        "client_id": st.secrets["GOOGLE_CLIENT_ID"],
        "client_secret": st.secrets["GOOGLE_CLIENT_SECRET"],
        "auth_uri": st.secrets.get("GOOGLE_AUTH_URI", "https://accounts.google.com/o/oauth2/auth"),
        "token_uri": st.secrets.get("GOOGLE_TOKEN_URI", "https://oauth2.googleapis.com/token"),
        "redirect_uris": [GOOGLE_REDIRECT_URI]
    }
}

# ================== SESSION STATE ==================
if "google_creds" not in st.session_state:
    st.session_state.google_creds = None

# ================== CALLBACK GOOGLE ==================
params = st.experimental_get_query_params()

if "code" in params and "google_redirect" in params and not st.session_state.google_creds:
    code = params["code"][0]

    flow = Flow.from_client_config(
        google_client_config,
        scopes=GOOGLE_SCOPES,
        redirect_uri=GOOGLE_REDIRECT_URI
    )
    flow.fetch_token(code=code)
    creds = flow.credentials

    st.session_state.google_creds = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes
    }

    st.experimental_set_query_params()
    st.success("✅ Autenticado con Google")
    st.experimental_rerun()

# ================== LOGIN GOOGLE ==================
st.subheader("🔐 Conexión con Google")

if st.session_state.google_creds:
    st.success("Conectado con Google")
else:
    if st.button("Conectar con Google"):
        flow = Flow.from_client_config(
            google_client_config,
            scopes=GOOGLE_SCOPES,
            redirect_uri=GOOGLE_REDIRECT_URI
        )
        auth_url, _ = flow.authorization_url(
            access_type="offline",
            include_granted_scopes="true",
            prompt="consent"
        )
        st.markdown(f"[Haz clic aquí para autorizar con Google]({auth_url})")

st.markdown("---")

# ================== ENVÍO DE EMAILS CON GMAIL ==================
st.header("📂 Envío masivo de correos")

uploaded_file = st.file_uploader("Sube un Excel con una columna 'email'", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.dataframe(df)

    if "email" not in df.columns:
        st.error("⚠️ El archivo debe tener una columna llamada 'email'")
    else:
        subject = st.text_input("Asunto del correo")
        body = st.text_area("Mensaje (puedes usar {nombre} etc.)")

        if st.button("📨 Enviar correos"):
            if not st.session_state.google_creds:
                st.error("Debes conectarte con Google primero.")
            else:
                try:
                    creds = Credentials(**st.session_state.google_creds)
                    service = build("gmail", "v1", credentials=creds)

                    total = len(df)
                    sent = 0
                    progress = st.progress(0)

                    for _, row in df.iterrows():
                        to = row["email"]

                        try:
                            personalized_body = body.format(**row.to_dict())
                        except:
                            personalized_body = body

                        raw_text = (
                            f"To: {to}\r\n"
                            f"Subject: {subject}\r\n\r\n"
                            f"{personalized_body}"
                        )

                        raw_bytes = base64.urlsafe_b64encode(raw_text.encode()).decode()
                        message = {"raw": raw_bytes}

                        service.users().messages().send(userId="me", body=message).execute()

                        sent += 1
                        progress.progress(sent / total)
                        st.info(f"Enviado a {to} ({sent}/{total})")

                    st.success("🎉 ¡Todos los correos fueron enviados!")

                except Exception as e:
                    st.error(f"Error al enviar correos: {e}")

else:
    st.info("Sube un archivo de Excel para comenzar.")
