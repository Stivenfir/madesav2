from __future__ import annotations

from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import base64
import json
import os
import smtplib
import subprocess
from typing import Iterable
from dotenv import load_dotenv

from RPAs_MADESA.Bases import Utils

load_dotenv()

# Lista de destinatarios del equipo de soporte (editar según necesidad)
DESTINATARIOS_MESA_AYUDA = [
    "mesaayuda@abcrepecev.com",
]

# Correo en tabla dbo.Login365 que se usa para autenticación OAuth/Graph
CORREO_365 = "mesaayuda@abcrepecev.com"

# Endpoints de Office 365
GRAPH_SENDMAIL_URL = "https://graph.microsoft.com/v1.0/me/sendMail"
SMTP_HOST = "smtp.office365.com"
SMTP_PORT = 587

# Archivo para evitar reenvíos duplicados del mismo pedido
PATH_ESTADO_NOTIFICACIONES = os.path.join(Utils.PATH_txt, "notificados_sin_guia.json")


def _bool_env(var_name: str, default: bool = False) -> bool:
    value = os.getenv(var_name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "si", "sí", "on"}


def _destinatarios_desde_env() -> list[str]:
    raw = os.getenv("NOTIFICAR_DESTINATARIOS", "").strip()
    if not raw:
        return DESTINATARIOS_MESA_AYUDA
    return [correo.strip() for correo in raw.split(",") if correo.strip()]


def _correo_origen_365() -> str:
    return os.getenv("NOTIFICAR_CORREO_365", CORREO_365).strip()


def _cargar_estado_notificados() -> set[str]:
    if not os.path.exists(PATH_ESTADO_NOTIFICACIONES):
        return set()
    with open(PATH_ESTADO_NOTIFICACIONES, "r", encoding="utf-8") as file:
        try:
            data = json.load(file)
        except json.JSONDecodeError:
            return set()
    return set(str(x) for x in data)


def _guardar_estado_notificados(pedidos: Iterable[str]) -> None:
    with open(PATH_ESTADO_NOTIFICACIONES, "w", encoding="utf-8") as file:
        json.dump(sorted(list(pedidos)), file, ensure_ascii=False, indent=2)


def _obtener_token_login365(correo: str) -> str:
    tablas_env = os.getenv("NOTIFICAR_TOKEN_TABLAS", "").strip()
    tablas_candidatas = (
        [t.strip() for t in tablas_env.split(",") if t.strip()]
        if tablas_env
        else ["dbo.Login365", "dbo.Login", "API.dbo.Login365", "API.dbo.Login"]
    )
    columnas_token = ["token", "access_token", "Token", "AccessToken"]
    columnas_correo = ["correo", "Correo", "email", "Email"]
    columnas_expira = ["expire_token", "Expire_token", "expires_on", "ExpiresOn"]

    errores: list[str] = []
    debug = _bool_env("NOTIFICAR_DEBUG", default=False)
    token_candidato = ""
    token_candidato_exp = 0
    ahora_epoch = int(datetime.now().timestamp())

    def _exp_token_jwt(token_jwt: str) -> int:
        try:
            partes = token_jwt.split(".")
            if len(partes) < 2:
                return 0
            payload_b64 = partes[1] + "=" * (-len(partes[1]) % 4)
            payload = json.loads(base64.urlsafe_b64decode(payload_b64.encode("utf-8")).decode("utf-8"))
            return int(payload.get("exp", 0))
        except Exception:
            return 0

    for tabla in tablas_candidatas:
        for col_token in columnas_token:
            for col_correo in columnas_correo:
                for col_expira in columnas_expira + [""]:
                    query = (
                        f"SELECT TOP 20 {col_token} FROM {tabla} WHERE {col_correo} = ? "
                        f"ORDER BY {col_expira} DESC"
                        if col_expira
                        else f"SELECT TOP 20 {col_token} FROM {tabla} WHERE {col_correo} = ?"
                    )
                    try:
                        Utils.cursor.execute(query, (correo,))
                        rows = Utils.cursor.fetchall()
                        for row in rows:
                            if not row or not row[0]:
                                continue
                            token_val = row[0]
                            exp_token = _exp_token_jwt(token_val)
                            if exp_token > ahora_epoch + 30:
                                if debug:
                                    print(
                                        f"[DEBUG] Token vigente tomado de {tabla}.{col_token} "
                                        f"filtrado por {col_correo}. exp={exp_token}"
                                    )
                                return token_val
                            if exp_token > token_candidato_exp:
                                token_candidato = token_val
                                token_candidato_exp = exp_token
                    except Exception as exc:
                        errores.append(f"{tabla}.{col_token}/{col_correo}/{col_expira or 'sin_orden'}: {exc}")

    if token_candidato:
        if debug:
            print(f"[DEBUG] No se encontró token vigente; se usa mejor candidato exp={token_candidato_exp}")
        return token_candidato

    raise RuntimeError(
        "No se encontró token en las tablas candidatas. "
        f"Correo: {correo}. Tablas probadas: {', '.join(tablas_candidatas)}"
    )


def _obtener_pedidos_sin_guia() -> list[dict]:
    Utils.cursorLite.execute(
        """
        SELECT id, idSharePoint, NumeroPedido, Nit, Transportadora, Ciudades, Departamento
        FROM InfoPedidos
        WHERE Guia IS NULL OR LENGTH(Guia) < 4
        ORDER BY id ASC
        """
    )
    rows = Utils.cursorLite.fetchall()
    headers = [col[0] for col in Utils.cursorLite.description]
    return [{headers[i]: row[i] for i in range(len(headers))} for row in rows]


def _obtener_pedidos_sin_guia_sharepoint() -> list[dict]:
    pedidos: list[dict] = []
    base_url = os.getenv("NOTIFICAR_SP_URL", "").strip() or Utils.V_Graph[3]
    url = f"{base_url}?expand=fields&$top=200"
    debug = _bool_env("NOTIFICAR_DEBUG", default=False)
    total_items = 0
    reintento_token_expirado = True

    while url:
        response = Utils.GraphPet(url, Utils.Cabecera)
        if "error" in response:
            error_code = (response.get("error", {}) or {}).get("code", "")
            error_message = ((response.get("error", {}) or {}).get("message", "") or "").lower()
            if (
                reintento_token_expirado
                and error_code == "InvalidAuthenticationToken"
                and "expired" in error_message
            ):
                _refrescar_token_graph_sharepoint()
                reintento_token_expirado = False
                continue
            raise RuntimeError(f"Error consultando SharePoint: {response['error']}")

        if "value" not in response:
            raise RuntimeError(f"Respuesta inesperada de SharePoint (sin 'value'): {response}")

        values = response.get("value", [])
        total_items += len(values)

        for item in values:
            fields = item.get("fields", {})
            numero_pedido = fields.get("Numero_pedido") or fields.get("ID") or item.get("id")
            guia = (
                fields.get("N_x002e_Guia")
                if fields.get("N_x002e_Guia") is not None
                else fields.get("Guia")
            )
            guia_txt = "" if guia is None else str(guia).strip()

            if not guia_txt or len(guia_txt) < 4 or guia is False:
                pedidos.append(
                    {
                        "id": item.get("id"),
                        "idSharePoint": item.get("id"),
                        "NumeroPedido": numero_pedido,
                        "Nit": fields.get("Nit"),
                        "Transportadora": fields.get("Transportadora"),
                        "Ciudades": fields.get("Ciudades"),
                        "Departamento": fields.get("Departamento"),
                    }
                )

        url = response.get("@odata.nextLink")

    if debug:
        print(f"[DEBUG] URL SharePoint usada: {base_url}")
        print(f"[DEBUG] Total items leídos en lista: {total_items}")
        print(f"[DEBUG] Total items sin guía detectados: {len(pedidos)}")

    return pedidos


def _refrescar_token_graph_sharepoint() -> str:
    comandos = [
        ["az.cmd", "account", "get-access-token", "--resource-type", "ms-graph"],
        ["az", "account", "get-access-token", "--resource-type", "ms-graph"],
    ]
    salida = None
    for cmd in comandos:
        try:
            salida = subprocess.run(cmd, capture_output=True, text=True, check=False)
            if salida.returncode == 0 and salida.stdout.strip():
                break
        except FileNotFoundError:
            continue

    if not salida or salida.returncode != 0:
        raise RuntimeError(
            "No fue posible refrescar el token de Graph. "
            "Valida Azure CLI y sesión activa con `az login`."
        )

    payload = Utils.loads(salida.stdout)
    token = payload.get("accessToken")
    if not token:
        raise RuntimeError("Azure CLI no devolvió accessToken para Graph.")

    Utils.Cabecera = {"Authorization": f"Bearer {token}"}
    try:
        Utils.cursorLite.execute("UPDATE SharePoint_CheckPoint SET token=? WHERE id=?", (token, Utils.V_Graph[0]))
        Utils.cnxnLite.commit()
    except Exception:
        # El script puede continuar aunque no sea posible persistir el token.
        pass
    return token


def _crear_html_notificacion(fallos: list[dict]) -> str:
    filas_html = ""
    for item in fallos:
        filas_html += (
            "<tr>"
            f"<td>{item.get('idSharePoint') or item.get('id') or ''}</td>"
            f"<td>{item['NumeroPedido']}</td>"
            f"<td>{item.get('Nit') or ''}</td>"
            f"<td>{item['Transportadora'] or ''}</td>"
            f"<td>{item['Ciudades'] or ''}</td>"
            f"<td>{item['Departamento'] or ''}</td>"
            "</tr>"
        )

    return f"""
    <html>
      <body>
        <p>Buen día,</p>
        <p>Se detectaron pedidos sin guía creada. Por favor validar:</p>
        <table border="1" cellpadding="4" cellspacing="0">
          <thead>
            <tr>
              <th>ID SharePoint</th>
              <th>Número Pedido</th>
              <th>NIT</th>
              <th>Transportadora</th>
              <th>Ciudad</th>
              <th>Municipio/Departamento</th>
            </tr>
          </thead>
          <tbody>
            {filas_html}
          </tbody>
        </table>
        <p>Fecha de ejecución: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
      </body>
    </html>
    """.strip()


def _enviar_por_graph(token: str, asunto: str, html: str, destinatarios: list[str]) -> None:
    to_recipients = [{"emailAddress": {"address": correo}} for correo in destinatarios]
    payload = {
        "message": {
            "subject": asunto,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": to_recipients,
        },
        "saveToSentItems": "true",
    }

    response = Utils.requests.post(
        GRAPH_SENDMAIL_URL,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )

    if response.status_code >= 300:
        raise RuntimeError(f"Graph sendMail falló [{response.status_code}]: {response.text}")


def _enviar_por_smtp_oauth(
    token: str,
    asunto: str,
    html: str,
    destinatarios: list[str],
    correo_origen: str,
) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = correo_origen
    msg["To"] = ", ".join(destinatarios)
    msg.attach(MIMEText(html, "html", "utf-8"))

    auth_string = f"user={correo_origen}\x01auth=Bearer {token}\x01\x01"
    auth_b64 = base64.b64encode(auth_string.encode("utf-8")).decode("utf-8")

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        code, response = server.docmd("AUTH", "XOAUTH2 " + auth_b64)
        if code != 235:
            raise RuntimeError(f"Error SMTP AUTH XOAUTH2 [{code}]: {response}")
        server.sendmail(correo_origen, destinatarios, msg.as_string())


def _enviar_por_smtp_password(
    asunto: str,
    html: str,
    destinatarios: list[str],
    correo_origen: str,
    clave_smtp: str,
) -> None:
    msg = MIMEMultipart("alternative")
    msg["Subject"] = asunto
    msg["From"] = correo_origen
    msg["To"] = ", ".join(destinatarios)
    msg.attach(MIMEText(html, "html", "utf-8"))

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(correo_origen, clave_smtp)
        server.sendmail(correo_origen, destinatarios, msg.as_string())


def main() -> None:
    modo_prueba = _bool_env("NOTIFICAR_MODO_PRUEBA", default=False)
    forzar_reenvio = _bool_env("NOTIFICAR_FORZAR_REENVIO", default=False)
    destinatarios = _destinatarios_desde_env()
    correo_origen = _correo_origen_365()
    clave_smtp = os.getenv("NOTIFICAR_SMTP_PASS", "").strip()
    fuente_datos = os.getenv("NOTIFICAR_FUENTE", "sqlite").strip().lower()
    if fuente_datos not in {"sqlite", "sharepoint"}:
        fuente_datos = "sqlite"

    print(f"Fuente de datos configurada: {fuente_datos}")

    if fuente_datos == "sharepoint":
        pedidos = _obtener_pedidos_sin_guia_sharepoint()
    else:
        pedidos = _obtener_pedidos_sin_guia()

    if not pedidos:
        print(f"No hay pedidos pendientes sin guía para notificar. Fuente: {fuente_datos}")
        return

    estado_notificados = _cargar_estado_notificados()
    pendientes = pedidos if forzar_reenvio else [p for p in pedidos if str(p["NumeroPedido"]) not in estado_notificados]

    if not pendientes:
        print("No hay pedidos nuevos sin guía para notificar (sin duplicados).")
        return

    if modo_prueba:
        print("=== MODO PRUEBA ACTIVADO ===")
        print(f"Destinatarios de prueba: {', '.join(destinatarios)}")
        print(f"Pedidos detectados sin guía: {len(pendientes)}")
        for pedido in pendientes[:10]:
            print(
                f"- ID SharePoint {pedido.get('idSharePoint') or pedido.get('id')} | "
                f"Pedido {pedido['NumeroPedido']} | NIT {pedido.get('Nit')} | "
                f"Transportadora: {pedido['Transportadora']} | "
                f"Ciudad: {pedido['Ciudades']} | "
                f"Municipio/Departamento: {pedido['Departamento']}"
            )
        print("No se envió correo ni se actualizó estado de notificados por estar en modo prueba.")
        return

    asunto = f"[ALERTA] Pedidos sin guía ({len(pendientes)})"
    html = _crear_html_notificacion(pendientes)

    if clave_smtp:
        _enviar_por_smtp_password(asunto, html, destinatarios, correo_origen, clave_smtp)
        canal = "SMTP Usuario/Clave"
    else:
        token = _obtener_token_login365(correo_origen)

        try:
            _enviar_por_graph(token, asunto, html, destinatarios)
            canal = "Graph API"
        except Exception as graph_error:
            print(f"Fallo envío por Graph API: {graph_error}")
            graph_error_txt = str(graph_error).lower()
            if "invalidauthenticationtoken" in graph_error_txt and "expired" in graph_error_txt:
                try:
                    token = _refrescar_token_graph_sharepoint()
                except Exception as refresh_error:
                    print(f"Reintento Graph con token refrescado falló: {refresh_error}")
                    print("Intentando nuevo token desde base de datos...")
                    token = _obtener_token_login365(correo_origen)

                try:
                    _enviar_por_graph(token, asunto, html, destinatarios)
                    canal = "Graph API (token actualizado)"
                except Exception as graph_retry_error:
                    print(f"Reintento Graph con token actualizado falló: {graph_retry_error}")
                    _enviar_por_smtp_oauth(token, asunto, html, destinatarios, correo_origen)
                    canal = "SMTP OAuth2"
            else:
                _enviar_por_smtp_oauth(token, asunto, html, destinatarios, correo_origen)
                canal = "SMTP OAuth2"

    for pedido in pendientes:
        estado_notificados.add(str(pedido["NumeroPedido"]))
    _guardar_estado_notificados(estado_notificados)

    print(
        f"Notificación enviada por {canal}. "
        f"Pedidos notificados: {len(pendientes)}. Destinatarios: {', '.join(destinatarios)}"
    )


if __name__ == "__main__":
    main()