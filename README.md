# Registro de Actividades (Streamlit)

App simple para que los especialistas registren actividades **por día** y marquen el resultado:
- **✓** Cumplido
- **✗** Incumplido

Incluye:
- Registro por fecha única o rango (con filtro por días de semana)
- Tablero mensual (carga y cumplimiento por especialista)
- Exportación a matriz mensual tipo Excel
- Modo **Administrador** (con código) para **editar el plazo/fecha** o borrar registros, con auditoría

## Cómo ejecutar (local)

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate

pip install -r requirements.txt
streamlit run app.py
```

## Código admin (secreto)

Crea el archivo:

`.streamlit/secrets.toml`

```toml
ADMIN_CODE = "TU_CODIGO_SUPER_SECRETO"
```

> **No** subas `secrets.toml` al repo (ya está en `.gitignore`).

También puedes usar variable de entorno:

```bash
export ADMIN_CODE="TU_CODIGO_SUPER_SECRETO"
```

## Persistencia de datos

Por defecto usa **SQLite** en `data/app.db`.
- En servidor propio / PC funciona perfecto.
- En Streamlit Community Cloud la persistencia puede reiniciarse si la app se redeploya. Para persistencia “real” en la nube, conecta a Postgres/Supabase o similar.

## Estructura

- `app.py` interfaz Streamlit
- `db.py` base de datos SQLite + exportación Excel
- `data/app.db` (se crea automáticamente)
