import streamlit as st
import pandas as pd
import base64
import requests

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import msal

st.set_page_config(page_title="Correo masivo OAuth", page_icon="📧")

st.title("📧 Envío automático de correos (Google / Outlook)")

# ================== CONFIG ==================

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
GOOGLE_REDIRECT_URI = st.secrets["GOOGLE_REDIRECT_URI"]

MS_CLIENT_ID = st.secrets["MS_CLIENT_ID"]
MS_TENANT_ID = st.secrets["MS_TENANT_ID"]
MS_CLIENT_SECRET = st.secrets["MS_CLIENT_SECRET"]
MS_REDIRECT_URI = st.secrets["MS_REDIRECT_URI"]
MS_SCOPES = ["https://graph.microsoft.com/Mail.Send"]

# Config de Google a partir de secrets (en vez de JSON)
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

if "ms_token" not in st.session_state:
    st.session_state.ms_token = None

# ================== MANEJO DE CALLBACK (query params) ==================

params = st.experimental_get_query_params()

# Callback de Google
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

    # Limpiar parámetros para no procesar el callback en cada recarga
    st.experimental_set_query_params()
    st.success("✅ Autenticado con Google")
    st.experimental_rerun()

# Callback de Microsoft
if "code" in params and "ms_auth" in params and not st.session_state.ms_token:
    code = params["code"][0]

    ms_app = msal.ConfidentialClientApplication(
        MS_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{MS_TENANT_ID}",
        client_credential=MS_CLIENT_SECRET
    )

    result = ms_app.acquire_token_by_authorization_code(
        code,
        scopes=MS_SCOPES,
        redirect_uri=MS_REDIRECT_URI
    )

    if "access_token" in result:
        st.session_state.ms_token = result
        st.experimental_set_query_params()
        st.success("✅ Autenticado con Microsoft")
        st.experimental_rerun()
    else:
        st.error(f"Error autenticando con Microsoft: {result.get('error_description')}")

# ================== SECCIÓN LOGIN ==================

st.subheader("🔐 Conexión con proveedores de correo")

# ---- GOOGLE LOGIN ----
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 🔵 Google (Gmail)")
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

# ---- MICROSOFT LOGIN ----
with col2:
    st.markdown("### 🟣 Microsoft (Outlook / Office 365)")
    if st.session_state.ms_token:
        st.success("Conectado con Microsoft")
    else:
        if st.button("Conectar con Microsoft"):
            ms_app = msal.ConfidentialClientApplication(
                MS_CLIENT_ID,
                authority=f"https://login.microsoftonline.com/{MS_TENANT_ID}",
                client_credential=MS_CLIENT_SECRET
            )
            auth_url = ms_app.get_authorization_request_url(
                scopes=MS_SCOPES,
                redirect_uri=MS_REDIRECT_URI
            )
            st.markdown(f"[Haz clic aquí para autorizar con Microsoft]({auth_url})")

st.markdown("---")

# ================== SECCIÓN DE ENVÍO DE CORREOS ==================

st.header("📂 Envío masivo de correos")

uploaded_file = st.file_uploader("Sube un Excel con una columna 'email'", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.write("📋 Vista previa del archivo:")
    st.dataframe(df)

    if "email" not in df.columns:
        st.error("⚠️ El archivo debe tener una columna llamada 'email'")
    else:
        subject = st.text_input("Asunto del correo")
        body = st.text_area("Mensaje del correo (puedes usar {nombre}, {campo2}, etc.)")

        provider = st.selectbox("Enviar usando", ["Selecciona", "Google (Gmail)", "Microsoft (Outlook)"])

        if st.button("📨 Enviar correos"):
            if not subject or not body:
                st.error("Por favor ingresa asunto y mensaje.")
            elif provider == "Selecciona":
                st.error("Selecciona un proveedor para enviar.")
            elif provider == "Google (Gmail)" and not st.session_state.google_creds:
                st.error("Primero conecta tu cuenta de Google.")
            elif provider == "Microsoft (Outlook)" and not st.session_state.ms_token:
                st.error("Primero conecta tu cuenta de Microsoft.")
            else:
                try:
                    total = len(df)
                    sent = 0
                    progress = st.progress(0)

                    # ======== ENVÍO CON GOOGLE ========
                    if provider == "Google (Gmail)":
                        creds = Credentials(**st.session_state.google_creds)
                        service = build("gmail", "v1", credentials=creds)

                        for _, row in df.iterrows():
                            to = row["email"]
                            try:
                                personalized_body = body.format(**row.to_dict())
                            except Exception:
                                personalized_body = body  # por si faltan columnas

                            raw_text = (
                                f"To: {to}\r\n"
                                f"Subject: {subject}\r\n"
                                f"\r\n"
                                f"{personalized_body}"
                            )

                            raw_bytes = base64.urlsafe_b64encode(raw_text.encode("utf-8")).decode("utf-8")
                            message = {"raw": raw_bytes}

                            service.users().messages().send(userId="me", body=message).execute()

                            sent += 1
                            progress.progress(int(sent / total * 100))
                            st.info(f"✅ Enviado a {to} ({sent}/{total})")

                    # ======== ENVÍO CON MICROSOFT ========
                    elif provider == "Microsoft (Outlook)":
                        access_token = st.session_state.ms_token["access_token"]
                        headers = {
                            "Authorization": f"Bearer {access_token}",
                            "Content-Type": "application/json"
                        }
                        url = "https://graph.microsoft.com/v1.0/me/sendMail"

                        for _, row in df.iterrows():
                            to = row["email"]
                            try:
                                personalized_body = body.format(**row.to_dict())
                            except Exception:
                                personalized_body = body

                            data = {
                                "message": {
                                    "subject": subject,
                                    "body": {
                                        "contentType": "Text",
                                        "content": personalized_body
                                    },
                                    "toRecipients": [
                                        {"emailAddress": {"address": to}}
                                    ]
                                },
                                "saveToSentItems": True
                            }

                            requests.post(url, headers=headers, json=data)

                            sent += 1
                            progress.progress(int(sent / total * 100))
                            st.info(f"✅ Enviado a {to} ({sent}/{total})")

                    st.success("🎉 ¡Todos los correos fueron enviados!")

                except Exception as e:
                    st.error(f"❌ Error al enviar correos: {e}")
else:
    st.info("Sube un archivo de Excel para comenzar.")
