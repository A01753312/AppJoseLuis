import streamlit as st
import pandas as pd
import os
import base64
import requests
from urllib.parse import urlencode

from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import msal

st.set_page_config(page_title="Mailer OAuth", page_icon="📧")

GOOGLE_CLIENT_SECRET_FILE = "google_client_secret.json"
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
GOOGLE_REDIRECT_URI = "https://organic-memory-7v9557gqxxv52q4j-8501.app.github.dev/?google_redirect=1"



MS_CLIENT_ID = "TU_CLIENT_ID"
MS_TENANT_ID = "TU_TENANT_ID"
MS_CLIENT_SECRET = "TU_CLIENT_SECRET"
MS_REDIRECT_URI = "http://localhost:8501/ms-auth"
MS_SCOPES = ["https://graph.microsoft.com/Mail.Send"]

if "google_creds" not in st.session_state:
    st.session_state.google_creds = None

if "ms_token" not in st.session_state:
    st.session_state.ms_token = None

st.title("📧 Envío de correos con Google / Microsoft")

# ========== LOGIN GOOGLE ==========
st.subheader("🔵 Iniciar sesión con Google")

if not st.session_state.google_creds:
    if st.button("Login con Google"):
        flow = Flow.from_client_secrets_file(
            GOOGLE_CLIENT_SECRET_FILE,
            scopes=GOOGLE_SCOPES,
            redirect_uri=GOOGLE_REDIRECT_URI
        )
        auth_url, state = flow.authorization_url(prompt="consent")
        st.experimental_set_query_params(google_state=state)
        st.write(f"[Haz clic aquí para autorizar con Google]({auth_url})")
else:
    st.success("✅ Conectado con Google")

# Procesar callback de Google
params = st.experimental_get_query_params()
if "code" in params and "google_state" in params:
    code = params["code"][0]
    state = params["google_state"][0]

    flow = Flow.from_client_secrets_file(
        GOOGLE_CLIENT_SECRET_FILE,
        scopes=GOOGLE_SCOPES,
        redirect_uri=GOOGLE_REDIRECT_URI,
        state=state
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
    st.experimental_rerun()

# ========== LOGIN MICROSOFT ==========
st.subheader("🟣 Iniciar sesión con Microsoft")

ms_app = msal.ConfidentialClientApplication(
    MS_CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{MS_TENANT_ID}",
    client_credential=MS_CLIENT_SECRET
)

if not st.session_state.ms_token:
    if st.button("Login con Microsoft"):
        auth_url = ms_app.get_authorization_request_url(
            scopes=MS_SCOPES,
            redirect_uri=MS_REDIRECT_URI
        )
        st.write(f"[Haz clic aquí para autorizar con Microsoft]({auth_url})")
else:
    st.success("✅ Conectado con Microsoft")

# Procesar callback de Microsoft (code en query params)
if "code" in params and "msauth" in params:
    code = params["code"][0]
    result = ms_app.acquire_token_by_authorization_code(
        code,
        scopes=MS_SCOPES,
        redirect_uri=MS_REDIRECT_URI
    )
    if "access_token" in result:
        st.session_state.ms_token = result
        st.experimental_rerun()
    else:
        st.error(f"Error autenticando con Microsoft: {result.get('error_description')}")

# ========== UI PARA ENVIAR CORREOS ==========
st.header("📂 Enviar correos masivos")

uploaded_file = st.file_uploader("Sube un Excel con una columna 'email'", type=["xlsx", "xls"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.dataframe(df)

    if "email" not in df.columns:
        st.error("El archivo debe tener una columna 'email'")
    else:
        subject = st.text_input("Asunto")
        body = st.text_area("Mensaje (puedes usar {nombre}, etc.)")
        provider = st.selectbox("Enviar usando", ["Google", "Microsoft"])

        if st.button("Enviar correos"):
            if provider == "Google" and st.session_state.google_creds:
                creds = Credentials(**st.session_state.google_creds)
                service = build("gmail", "v1", credentials=creds)
                for _, row in df.iterrows():
                    to = row["email"]
                    text = body.format(**row.to_dict())
                    raw_msg = f"To: {to}\r\nSubject: {subject}\r\n\r\n{txt}"
                    message = {
                        "raw": base64.urlsafe_b64encode(raw_msg.encode()).decode()
                    }
                    service.users().messages().send(userId="me", body=message).execute()
                st.success("✅ Correos enviados con Google")

            elif provider == "Microsoft" and st.session_state.ms_token:
                access_token = st.session_state.ms_token["access_token"]
                headers = {"Authorization": f"Bearer {access_token}",
                           "Content-Type": "application/json"}
                url = "https://graph.microsoft.com/v1.0/me/sendMail"

                for _, row in df.iterrows():
                    to = row["email"]
                    text = body.format(**row.to_dict())
                    data = {
                        "message": {
                            "subject": subject,
                            "body": {
                                "contentType": "Text",
                                "content": text
                            },
                            "toRecipients": [
                                {"emailAddress": {"address": to}}
                            ]
                        },
                        "saveToSentItems": "true"
                    }
                    requests.post(url, headers=headers, json=data)
                st.success("✅ Correos enviados con Microsoft")
            else:
                st.error("Debes iniciar sesión con el proveedor elegido.")
