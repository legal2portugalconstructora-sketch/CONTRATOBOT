CONTRATO — Proyecto de automatización Excel/Word

Descripción
- La app procesa hojas Excel según reglas definidas (copia valores, aplica
  formatos y genera alertas) y genera/edita el CONTRATO en Word. Usa
  Streamlit como interfaz, organizada como una app multipágina:
  - `Inicio.py` — punto de entrada.
  - `pages/1_🤖CONTRATO_Excel.py` — Pasos 1 a 6 sobre los archivos Excel.
  - `pages/2_🤖CONTRATO_Word.py` — generación del CONTRATO en Word.
  - `excel_pipeline.py`, `word_pipeline.py`, `word_com_bridge.py`,
    `ui_utils.py` — lógica compartida importada por esas páginas.
  - `auth_utils.py` — login compartido; ver más abajo.
- Es una app **interna** de acceso restringido: ninguna página se ve ni se
  usa sin iniciar sesión.

Archivos importantes
- [Inicio.py](Inicio.py)
- [auth_utils.py](auth_utils.py)
- [excel_pipeline.py](excel_pipeline.py)
- [word_pipeline.py](word_pipeline.py)
- [word_com_bridge.py](word_com_bridge.py)
- memoria_aprendizaje.json

## 1. Instalar dependencias

Requiere Python 3.10+ instalado en la máquina que va a correr el servidor
(normalmente una PC con Excel/Windows, porque `word_com_bridge.py` usa
automatización COM de Excel).

```powershell
pip install -r requirements.txt
```

## 2. Configurar usuarios (`config.yaml`)

`config.yaml` **no** está en el repositorio (está en `.gitignore`) porque
contiene las contraseñas hasheadas de cada usuario. La primera vez:

1. Copiá `config.example.yaml` a `config.yaml`.
2. Para cada usuario, generá el hash de su contraseña. Desde una consola
   de Python, con el entorno ya con `streamlit-authenticator` instalado:

   ```python
   import streamlit_authenticator as stauth
   print(stauth.Hasher().hash("la-contraseña-en-texto-plano"))
   ```

   Eso imprime un hash bcrypt (empieza con `$2b$...`). Pegá ese hash en el
   campo `password` del usuario correspondiente en `config.yaml`. **Nunca**
   escribas la contraseña en texto plano en el archivo.

3. Estructura de cada usuario dentro de `credentials.usernames`:

   ```yaml
   credentials:
     usernames:
       nombre_de_usuario:
         name: Nombre visible
         email: correo@empresa.com
         password: $2b$12$...hash...
   ```

4. Para agregar o quitar un usuario, sumá o borrá su bloque bajo
   `usernames`. Para cambiar una contraseña, repetí el paso 2 y reemplazá
   solo el `password` de ese usuario.

5. `cookie.key` es la clave con la que se firma la cookie de sesión.
   Generala una sola vez (no hace falta volver a tocarla salvo que quieras
   invalidar todas las sesiones activas):

   ```python
   import secrets
   print(secrets.token_hex(32))
   ```

### Usuarios de ejemplo ya incluidos

`config.yaml` viene con 3 usuarios de ejemplo listos para el primer
ingreso. **Cambiá estas contraseñas apenas puedas** (repitiendo el paso 2
de arriba) — quedan documentadas acá solo porque no son recuperables desde
el hash:

| Usuario  | Contraseña      |
|----------|-----------------|
| admin    | `9n06P&zTzv7i!7` |
| asesor1  | `exMj3AN73OoAmU` |
| asesor2  | `LTR0QW*4fJja#k` |

## 3. Variables sensibles (`.env`)

Cualquier valor sensible que no sea de usuarios (API keys, etc.) va en
`.env` (también excluido de `.gitignore`), en formato `CLAVE=valor`. El
código lo lee con `python-dotenv` (`load_dotenv()` en `Inicio.py`).

## 4. Levantar el servidor en la red interna

Desde la carpeta del proyecto:

```powershell
streamlit run Inicio.py --server.address 0.0.0.0 --server.port 8501
```

- `--server.address 0.0.0.0` hace que el servidor escuche en todas las
  interfaces de red, no solo `localhost`, para que otras computadoras de
  la red interna puedan acceder.
- El puerto por defecto de Streamlit es **8501**. Para usar otro puerto,
  cambiá el valor de `--server.port` (por ejemplo `8502`).
- Desde otra computadora de la misma red, entrá a
  `http://IP-DE-ESTA-PC:8501` (reemplazando `IP-DE-ESTA-PC` por la IP
  local de la máquina que corre el servidor, ej. `192.168.1.50`).
- Esta app **no** está pensada para exponerse a internet directamente
  (sin HTTPS/reverse proxy); es para uso dentro de la red interna de la
  empresa.

## 5. Cómo funciona el login

- Todas las páginas (`Inicio.py` y las de `pages/`) exigen sesión iniciada
  — Streamlit no comparte el login entre páginas automáticamente, así que
  cada una llama a `auth_utils.requerir_login()` al principio.
- Mientras no haya sesión iniciada, la página no muestra ningún contenido
  propio, solo el formulario de login.
- El botón **"Cerrar sesión"** está siempre visible en la barra lateral.
- La sesión se mantiene por cookie durante 7 días (configurable en
  `config.yaml`, `cookie.expiry_days`).
