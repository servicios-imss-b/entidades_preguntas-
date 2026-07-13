import pandas as pd
import os 
usuario = os.getlogin() 
base = pd.read_excel(fr"C:\Users\{usuario}\Downloads\nuevo-f\formulario-ang\bases\UnidadesIMB_CS!_v2.xlsx",sheet_name="Sheet 1")
clues = pd.read_parquet(fr"C:\Users\{usuario}\IMSS-BIENESTAR\División de Procesamiento de información - Repositorio de Datos\CLUES\clues.parquet")
import pandas as pd

sheet_id = "1maRNGDuU9rEFWZLgMdhJS1waAnJxl6ENntm-nyD0tq8"
gid = "1765182479"

url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"

base_an = pd.read_csv(url)
col = ["clues_imb"]
base = base[col]
base = base.merge(
    clues[["clues_imb", "entidad",'nombre_de_la_unidad']],
    on="clues_imb",
    how="left"
)
colum =['clues_imb', 'entidad','consultorio','pregunta','nombre_de_la_unidad']
base_an = base_an[colum]    
b = pd.read_excel(fr"C:\Users\{usuario}\Downloads\nuevo-f\formulario-ang\bases\UM_IMB_SUS.xlsx",sheet_name="Hoja2")
b = b.drop(index=[0, 2, 5,3,1])
b = b.drop(columns=['Unnamed: 1'])
b = b.rename(columns={
   'CLUES' : 'preguntas',
})
base_an = base_an[
    ~base_an["pregunta"].str.contains("internet|turno_consultorio|consultorios_habilitados", case=False, na=False)
]
base_conteo =base["clues_imb"].unique()
base_conteo
total_unidades = base["clues_imb"].nunique()

respondieron = base_an["clues_imb"].nunique()

sin_responder = total_unidades - respondieron
base_an = base_an.drop_duplicates()
avance_general = round(respondieron / total_unidades * 100, 1)
# CLUES pendientes: unidades que no han enviado ninguna respuesta
clues_catalogo = (
    base[["clues_imb", "entidad", "nombre_de_la_unidad"]]
    .drop_duplicates()
    .copy()
)

clues_respondieron = base_an[["clues_imb"]].drop_duplicates()

clues_pendientes = (
    clues_catalogo
    .merge(clues_respondieron, on="clues_imb", how="left", indicator=True)
    .query("_merge == 'left_only'")
    .drop(columns=["_merge"])
    .sort_values(["entidad", "nombre_de_la_unidad", "clues_imb"])
    .reset_index(drop=True)
)

print(f"CLUES pendientes (sin respuestas): {len(clues_pendientes)}")
# TABLA 2 - COMPLETITUD POR UNIDAD

# Número de preguntas del formulario
n_preguntas = len(b)

# Universo de unidades (todas las CLUES del catálogo base)
universo_unidades = (
    base[["clues_imb", "entidad", "nombre_de_la_unidad"]]
    .drop_duplicates()
    .copy()
)

# Número de consultorios por unidad (solo de las que ya respondieron)
consultorios = (
    base_an
    .groupby(["clues_imb", "entidad", "nombre_de_la_unidad"], as_index=False)["consultorio"]
    .max()
    .rename(columns={"consultorio": "consultorios"})
)

# Preguntas respondidas por unidad
respondidas = (
    base_an
    .groupby(["clues_imb", "entidad", "nombre_de_la_unidad"])
    .size()
    .reset_index(name="respondidas")
)

# Tabla final incluyendo unidades sin respuestas
tabla_unidades = (
    universo_unidades
    .merge(consultorios, on=["clues_imb", "entidad", "nombre_de_la_unidad"], how="left")
    .merge(respondidas, on=["clues_imb", "entidad", "nombre_de_la_unidad"], how="left")
)

tabla_unidades[["consultorios", "respondidas"]] = (
    tabla_unidades[["consultorios", "respondidas"]]
    .fillna(0)
    .astype(int)
)

tabla_unidades["esperadas"] = tabla_unidades["consultorios"] * n_preguntas

tabla_unidades["porcentaje"] = 0.0
mask = tabla_unidades["esperadas"] > 0
tabla_unidades.loc[mask, "porcentaje"] = (
    tabla_unidades.loc[mask, "respondidas"]
    / tabla_unidades.loc[mask, "esperadas"]
    * 100
).round(1)

tabla_unidades = (
    tabla_unidades
    .rename(columns={"clues_imb": "clues"})
    .sort_values(["porcentaje", "entidad", "nombre_de_la_unidad"])
    .reset_index(drop=True)
)
import json
from pathlib import Path
from datetime import datetime

# Configuracion de colores (como tu front)
COLOR_PRIMARIO = "#FAF2F5"
COLOR_SECUNDARIO = "#AE8640"
COLOR_HBC = "#FDFDFDC0"
COLOR_FONDO = "#235B4E"
COLOR_BORDE = "#7A1737"
COLOR_TEXTO = "#000000"

# CLUES pendientes por regla solicitada
pendientes = (
    tabla_unidades[tabla_unidades["porcentaje"] < 100]
    .copy()
    .sort_values(["porcentaje", "entidad", "clues"], ascending=[False, True, True])
    .reset_index(drop=True)
)

if pendientes.empty:
    print("No hay CLUES pendientes (todas las unidades estan al 100%).")
else:
    resumen_entidad = (
        pendientes.groupby("entidad", as_index=False)
        .agg(clues_pendientes=("clues", "count"), avance_promedio=("porcentaje", "mean"))
        .sort_values(["entidad"], ascending=[True])
    )

    fecha_reporte = datetime.now().strftime("%Y-%m-%d %H:%M")

    # Datos para JS por entidad (orden inicial: mayor a menor porcentaje)
    datos_entidades = {}
    for entidad, sub in pendientes.groupby("entidad", sort=False):
        sub = sub.sort_values(["porcentaje", "clues"], ascending=[False, True])
        datos_entidades[entidad] = [
            {
                "clues": str(r["clues"]),
                "nombre_de_la_unidad": str(r["nombre_de_la_unidad"]),
                "consultorios": int(r["consultorios"]),
                "respondidas": int(r["respondidas"]),
                "esperadas": int(r["esperadas"]),
                "porcentaje": float(r["porcentaje"]),
            }
            for _, r in sub.iterrows()
        ]

    botones_entidad = []
    for _, r in resumen_entidad.iterrows():
        entidad = r["entidad"]
        botones_entidad.append(
            f"""
            <button class='estado-btn' onclick='abrirEntidad("{entidad.replace("\"", "\\\"")}")'>
                {entidad}
            </button>
            """
        )

    payload_js = {
        "fecha_reporte": fecha_reporte,
        "datos_entidades": datos_entidades,
    }

    html_content = f"""<!DOCTYPE html>
<html lang='es'>
<head>
    <meta charset='UTF-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>Informe de CLUES Pendientes - IMSS Bienestar</title>
    <link href='https://fonts.googleapis.com/css2?family=League+Spartan:wght@400;500;600;700&display=swap' rel='stylesheet'>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; font-family: 'League Spartan', sans-serif; }}
        body {{ background: linear-gradient(135deg, {COLOR_PRIMARIO} 0%, #fff 100%); min-height: 100vh; color: {COLOR_TEXTO}; padding: 20px; }}
        .container {{ max-width: 1400px; margin: 0 auto; }}

        .menu-container {{ position: fixed; top: 20px; right: 20px; z-index: 1001; }}
        .menu-btn {{ background: {COLOR_FONDO}; border: none; border-radius: 50%; width: 50px; height: 50px; cursor: pointer; display: flex; flex-direction: column; justify-content: center; align-items: center; gap: 6px; box-shadow: 0 2px 10px rgba(0,0,0,.2); transition: .3s; }}
        .menu-btn:hover {{ background: {COLOR_SECUNDARIO}; transform: scale(1.05); }}
        .menu-btn span {{ width: 25px; height: 3px; background: #fff; border-radius: 3px; }}

        .menu-overlay {{ position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 1001; display: none; }}
        .menu-overlay.active {{ display: block; }}

        .menu-panel {{ position: fixed; top: 0; right: -420px; width: 390px; height: 100%; background: #fff; box-shadow: -2px 0 10px rgba(0,0,0,.1); z-index: 1002; transition: .3s; overflow-y: auto; padding: 80px 24px 24px; }}
        .menu-panel.active {{ right: 0; }}
        .close-menu {{ position: absolute; top: 18px; right: 18px; background: none; border: none; font-size: 30px; color: {COLOR_FONDO}; cursor: pointer; }}
        .menu-panel h2 {{ color: {COLOR_FONDO}; margin-bottom: 14px; border-bottom: 3px solid {COLOR_SECUNDARIO}; padding-bottom: 8px; }}
        .menu-panel h3 {{ color: {COLOR_SECUNDARIO}; margin: 14px 0 8px; }}
        .menu-panel p, .menu-panel li {{ color: #444; line-height: 1.5; }}
        .menu-panel ul {{ margin-left: 20px; }}

        .header {{ background: {COLOR_FONDO}; color: #fff; border-radius: 15px; text-align: center; padding: 20px; margin-bottom: 20px; }}
        .header img {{ height: 50px; margin-bottom: 10px; }}
        .header h1 {{ font-size: 28px; margin-bottom: 6px; }}

        .estados-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 14px; }}
        .estado-btn {{ background: {COLOR_FONDO}; color: #fff; border: none; border-radius: 12px; padding: 14px 10px; font-size: 14px; font-weight: 600; cursor: pointer; transition: .3s; }}
        .estado-btn:hover {{ background: {COLOR_SECUNDARIO}; transform: translateY(-2px); }}

        .modal {{ display: none; position: fixed; inset: 0; background: rgba(0,0,0,.65); z-index: 2000; align-items: center; justify-content: center; }}
        .modal.active {{ display: flex; }}
        .modal-content {{ width: 96%; max-width: 1320px; max-height: 90vh; overflow: auto; background: #fff; border-radius: 16px; padding: 20px; position: relative; }}
        .close-btn {{ position: absolute; top: 10px; right: 14px; border: none; background: none; font-size: 30px; cursor: pointer; color: #333; }}
        .modal h2 {{ color: {COLOR_FONDO}; margin-bottom: 8px; }}
        .modal .sub {{ margin-bottom: 12px; color: #444; }}

        .sort-actions {{ display: flex; gap: 10px; margin-bottom: 12px; flex-wrap: wrap; }}
        .sort-btn {{ background: {COLOR_SECUNDARIO}; color: #fff; border: none; border-radius: 8px; padding: 8px 12px; cursor: pointer; font-weight: 600; }}
        .sort-btn:hover {{ filter: brightness(0.95); }}

        .table-wrap {{ max-height: 65vh; overflow: auto; border: 1px solid #ececec; border-radius: 10px; }}
        table {{ width: 100%; border-collapse: collapse; min-width: 980px; }}
        th, td {{ border: 1px solid #eaeaea; padding: 10px 12px; text-align: left; }}
        th {{ background: {COLOR_FONDO}; color: #fff; position: sticky; top: 0; z-index: 1; }}
        tbody tr:nth-child(even) {{ background: #fafafa; }}
        tbody tr:hover {{ background: #fff3e0; }}

        .footer {{ margin-top: 18px; color: #666; font-size: 14px; text-align: center; }}
    </style>
</head>
<body>
    <div class='menu-container'>
        <button class='menu-btn' onclick='toggleMenu()'>
            <span></span><span></span><span></span>
        </button>
    </div>
    <div class='menu-overlay' id='menuOverlay' onclick='toggleMenu()'></div>
    <aside class='menu-panel' id='menuPanel'>
        <button class='close-menu' onclick='toggleMenu()'>&times;</button>
        <h2>Instrucciones</h2>
        <h3>Objetivo</h3>
        <p>Ver unidades pendientes por entidad.</p>
        <h3>Como usarlo</h3>
        <ul>
            <li>Haz clic en una entidad para abrir su tabla.</li>
            <li>Usa los botones para ordenar por porcentaje.</li>
        </ul>
        <h3>Fecha de corte</h3>
        <p>{fecha_reporte}</p>
    </aside>

    <div class='container'>
        <header class='header'>
            <img src='https://imssbienestar.gob.mx/assets/img/imb_b.svg' alt='IMSS Bienestar'>
            <h1>INFORME DE CLUES PENDIENTES POR ENTIDAD</h1>
            <p>Actualizado: {fecha_reporte}</p>
        </header>

        <section class='estados-grid'>
            {''.join(botones_entidad)}
        </section>

        <p class='footer'>.</p>
    </div>

    <div class='modal' id='detalleModal'>
        <div class='modal-content'>
            <button class='close-btn' onclick='cerrarModal()'>&times;</button>
            <h2 id='tituloEntidad'>Entidad</h2>
            <p class='sub' id='subEntidad'></p>

            <div class='sort-actions'>
                <button class='sort-btn' onclick='ordenarPorPorcentajeDesc()'>Ordenar % mayor a menor</button>
                <button class='sort-btn' onclick='ordenarPorPorcentajeAsc()'>Ordenar % menor a mayor</button>
                <button class='sort-btn' onclick='ordenarPorClues()'>Ordenar por CLUES</button>
            </div>

            <div class='table-wrap'>
                <table>
                    <thead>
                        <tr>
                            <th>CLUES</th>
                            <th>Unidad</th>
                            <th>Consultorios</th>
                            <th>Respondidas</th>
                            <th>Esperadas</th>
                            <th>Avance</th>
                        </tr>
                    </thead>
                    <tbody id='tbodyDetalle'></tbody>
                </table>
            </div>
        </div>
    </div>

    <script src='data.js'></script>
    <script>
        const datosEntidades = (window.INFORME_DATA && window.INFORME_DATA.datos_entidades) || {{}};
        let rowsActuales = [];

        function toggleMenu() {{
            document.getElementById('menuPanel').classList.toggle('active');
            document.getElementById('menuOverlay').classList.toggle('active');
        }}

        function renderTabla() {{
            const tbody = document.getElementById('tbodyDetalle');
            tbody.innerHTML = '';
            rowsActuales.forEach(r => {{
                const tr = document.createElement('tr');
                tr.innerHTML = `
                    <td>${{r.clues}}</td>
                    <td>${{r.nombre_de_la_unidad}}</td>
                    <td>${{r.consultorios}}</td>
                    <td>${{r.respondidas}}</td>
                    <td>${{r.esperadas}}</td>
                    <td>${{Number(r.porcentaje).toFixed(1)}}%</td>`;
                tbody.appendChild(tr);
            }});
        }}

        function abrirEntidad(entidad) {{
            rowsActuales = [...(datosEntidades[entidad] || [])];
            renderTabla();
            document.getElementById('tituloEntidad').textContent = entidad;
            document.getElementById('subEntidad').textContent = 'Detalle de unidades';
            document.getElementById('detalleModal').classList.add('active');
        }}

        function ordenarPorPorcentajeDesc() {{
            rowsActuales.sort((a, b) => b.porcentaje - a.porcentaje || a.clues.localeCompare(b.clues));
            renderTabla();
        }}

        function ordenarPorPorcentajeAsc() {{
            rowsActuales.sort((a, b) => a.porcentaje - b.porcentaje || a.clues.localeCompare(b.clues));
            renderTabla();
        }}

        function ordenarPorClues() {{
            rowsActuales.sort((a, b) => a.clues.localeCompare(b.clues));
            renderTabla();
        }}

        function cerrarModal() {{
            document.getElementById('detalleModal').classList.remove('active');
        }}

        document.getElementById('detalleModal').addEventListener('click', function(e) {{
            if (e.target === this) cerrarModal();
        }});

        document.addEventListener('keydown', function(e) {{
            if (e.key === 'Escape') {{
                cerrarModal();
                document.getElementById('menuPanel').classList.remove('active');
                document.getElementById('menuOverlay').classList.remove('active');
            }}
        }});
    </script>
</body>
</html>"""

    js_content = "window.INFORME_DATA = " + json.dumps(payload_js, ensure_ascii=False) + ";\n"

    salida_dir = Path(__file__).resolve().parent
    salida_html = salida_dir / "index.html"
    salida_js = salida_dir / "data.js"
    salida_html.write_text(html_content, encoding="utf-8")
    salida_js.write_text(js_content, encoding="utf-8")
    print(f"Reporte generado: {salida_html}")
    print(f"Datos generados: {salida_js}")
    webbrowser.open(salida_html.as_uri())

    resumen_entidad.head(10)
