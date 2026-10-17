import pandas as pd
import plotly.express as px
import plotly.io as pio
import os
import urllib.request
import json

def generar_dashboard_riesgos():
    archivo = 'bd_siniestralidad.xlsx'
    if not os.path.exists(archivo):
        print("Falta el archivo Excel. Ejecuta primero el generador de datos.")
        return

    df = pd.read_excel(archivo).dropna(how='all')
    
    # KPIs base
    total_polizas = len(df)
    riesgo_alto = len(df[df['Nivel_Riesgo'] == 'Alto'])
    riesgo_medio = len(df[df['Nivel_Riesgo'] == 'Medio'])
    riesgo_bajo = len(df[df['Nivel_Riesgo'] == 'Bajo'])
    
    suma_total = df['Monto_Asegurado'].sum()
    suma_promedio = df['Monto_Asegurado'].mean()
    
    color_bajo = "#C5CAE9"  
    color_medio = "#7E57C2" 
    color_alto = "#311B92"  
    mapa_colores = {'Bajo': color_bajo, 'Medio': color_medio, 'Alto': color_alto}

    # Procesar JSON para inyectarlo en JS (para filtros dinámicos)
    json_data = df.to_json(orient='records')

    # Componentes vacíos, se dibujarán con JS
    
    # 7. Plantilla HTML Responsiva 
    html_template = f"""<!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Dashboard | Concentración de Riesgos</title>
        
        <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
        <script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
        <script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.js"></script>
        <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.css">
        
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;500;700;900&display=swap');
            :root {{
                --bg-body: #EEF0F4; --bg-card: #FFFFFF; --header-bg: #1A1A24;
                --text-main: #2C303A; --accent-lila: #C5CAE9; --accent-morado: #7E57C2; --accent-dark: #311B92;
            }}
            * {{ box-sizing: border-box; }}
            body {{ font-family: 'Inter', sans-serif; background-color: var(--bg-body); color: var(--text-main); margin: 0; padding: 20px; }}
            .container {{ max-width: 1400px; margin: 0 auto; }}
            
            .header {{ background: var(--header-bg); color: white; padding: 25px 30px; border-radius: 12px; border-bottom: 4px solid var(--accent-morado); margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px; }}
            .header h1 {{ margin: 0; font-size: 1.8rem; font-weight: 900; letter-spacing: -0.5px; }}
            .header p {{ margin: 5px 0 0 0; color: var(--accent-lila); font-size: 0.95rem; }}
            
            .tab-nav {{ display: flex; background: rgba(255,255,255,0.1); padding: 5px; border-radius: 8px; flex-wrap: wrap; align-items:center; }}
            .tab-label {{ background: none; border: none; color: #fff; padding: 10px 20px; font-weight: 700; cursor: pointer; border-radius: 6px; transition: 0.3s; }}
            .tab-label.active {{ background: var(--accent-morado); color: white; }}
            
            .btn-reset {{ background: #D32F2F; color: white; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; cursor: pointer; margin-left: 15px; display:none; }}
            .btn-reset:hover {{ background: #B71C1C; }}

            .tab-content {{ display: none; animation: fadeIn 0.4s; }}
            
            .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; margin-bottom: 25px; }}
            .kpi-card {{ background: var(--bg-card); padding: 20px; border-radius: 12px; border-left: 6px solid var(--accent-dark); transition: 0.3s; }}
            .kpi-card.clickable {{ cursor: pointer; }}
            .kpi-card.clickable:hover {{ transform: translateY(-3px); box-shadow: 0 4px 15px rgba(0,0,0,0.1); }}
            .kpi-card.active-filter {{ background: #f0f0fa; border-width: 8px; }}

            .kpi-title {{ font-size: 0.8rem; text-transform: uppercase; font-weight: 700; color: #787E8F; letter-spacing: 0.5px; }}
            .kpi-val {{ font-size: 2rem; font-weight: 900; margin-top: 10px; color: var(--header-bg); line-height: 1; }}
            
            .main-grid {{ display: grid; grid-template-columns: 1fr; gap: 20px; }}
            @media (min-width: 1024px) {{ .main-grid {{ grid-template-columns: 3fr 2fr; }} }}
            
            .chart-card {{ background: var(--bg-card); padding: 25px; border-radius: 12px; overflow: hidden; }}
            .chart-title {{ margin: 0 0 15px 0; font-size: 1.1rem; border-bottom: 2px dashed var(--accent-lila); padding-bottom: 10px; color: var(--header-bg); }}
            
            .map-container {{ display: grid; grid-template-columns: 1fr; gap: 20px; }}
            @media (min-width: 1200px) {{ .map-container {{ grid-template-columns: 3fr 1fr; align-items: start; }} }}
            
            .min-table {{ width: 100%; text-align: left; border-collapse: collapse; }}
            .min-table th {{ border-bottom: 2px solid var(--accent-morado); padding: 10px 5px; color: var(--header-bg); font-size: 0.85rem; }}
            .min-table td {{ padding: 10px 5px; border-bottom: 1px solid #E4E7EB; font-size: 0.9rem; font-weight: 500; cursor: pointer; transition: 0.2s; }}
            .min-table tbody tr:hover td {{ background: #f5f5f5; color: var(--accent-morado); font-weight: bold; }}

            .table-wrapper {{ background: var(--bg-card); padding: 25px; border-radius: 12px; overflow-x: auto; }}
            table.dataTable thead th {{ background: var(--header-bg); color: white; border: none; font-size: 0.85rem; }}
            
            .js-plotly-plot, .plot-container {{ width: 100% !important; }}
            
            @keyframes fadeIn {{ from {{ opacity: 0; }} to {{ opacity: 1; }} }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div>
                    <h1>Inteligencia de Siniestralidad</h1>
                    <p>Monitoreo Actuarial de Concentración de Riesgos</p>
                </div>
                <div class="tab-nav">
                    <button class="tab-label active" onclick="openTab(event, 'tab-resumen')">📊 Resumen Espacial</button>
                    <button class="tab-label" onclick="openTab(event, 'tab-datos')">📋 Base de Pólizas</button>
                    <button id="btnResetFilters" class="btn-reset" onclick="resetAllFilters()">Limpiar Filtros ✖</button>
                </div>
            </div>
            
            <div id="tab-resumen" class="tab-content" style="display: block;">
                <div class="kpi-grid">
                    <div class="kpi-card">
                        <div class="kpi-title">Cartera Total</div>
                        <div class="kpi-val" id="kpi-total">0</div>
                    </div>
                    <div class="kpi-card" style="border-left-color: #4CAF50;">
                        <div class="kpi-title">Suma Asegurada Total</div>
                        <div class="kpi-val" id="kpi-suma" style="color: #4CAF50; font-size: 1.6rem;">$0</div>
                    </div>
                    <div class="kpi-card clickable" id="kpi-card-bajo" style="border-left-color: var(--accent-lila);" onclick="toggleRiskFilter('Bajo')">
                        <div class="kpi-title">Riesgo Bajo (Filtrar)</div>
                        <div class="kpi-val" id="kpi-bajo" style="color: var(--accent-lila);">0</div>
                    </div>
                    <div class="kpi-card clickable" id="kpi-card-medio" style="border-left-color: var(--accent-morado);" onclick="toggleRiskFilter('Medio')">
                        <div class="kpi-title">Riesgo Medio (Filtrar)</div>
                        <div class="kpi-val" id="kpi-medio" style="color: var(--accent-morado);">0</div>
                    </div>
                    <div class="kpi-card clickable" id="kpi-card-alto" style="border-left-color: var(--accent-dark);" onclick="toggleRiskFilter('Alto')">
                        <div class="kpi-title">Riesgo Alto (Filtrar)</div>
                        <div class="kpi-val" id="kpi-alto" style="color: var(--accent-dark);">0 <span style="font-size:1rem; color:#787E8F;">Pólizas</span></div>
                    </div>
                </div>

                <div class="main-grid">
                    <div class="chart-card">
                        <h3 class="chart-title">📍 Dispersión Geográfica (Click para filtrar Estado)</h3>
                        <div class="map-container">
                            <div id="plotly-mapa" style="height: 500px;"></div>
                            <div>
                                <h4 style="margin: 0 0 10px 0; color: var(--accent-dark); font-size: 0.9rem;">⚠️ Zonas Críticas (Click para filtrar)</h4>
                                <div id="top5-container"></div>
                            </div>
                        </div>
                    </div>
                    
                    <div style="display: flex; flex-direction: column; gap: 20px;">
                        <div class="chart-card">
                            <h3 class="chart-title">📈 Riesgo por Ramo (Click para filtrar Ramo)</h3>
                            <div id="plotly-bar" style="height: 350px;"></div>
                        </div>
                        <div class="chart-card">
                            <h3 class="chart-title">Composición de Cartera</h3>
                            <div id="plotly-pie" style="height: 350px;"></div>
                        </div>
                    </div>
                </div>
            </div>

            <div id="tab-datos" class="tab-content">
                <div class="table-wrapper">
                    <h3 class="chart-title">Desglose a Nivel Póliza</h3>
                    <table id="tablaPolizas" class="display w-full"></table>
                </div>
            </div>
        </div>

        <script>
            // Datos inyectados desde Python
            const rawData = {json_data};
            let geojsonMx = null;
            
            // Estado Global de Filtros
            let currentFilters = {{
                estado: null,
                ramo: null,
                riesgo: null
            }};
            
            let dataTableInstance = null;

            // Paleta
            const color_bajo = "#C5CAE9";
            const color_medio = "#7E57C2";
            const color_alto = "#311B92";
            
            // Formatters
            const fmtMoney = new Intl.NumberFormat('es-MX', {{ style: 'currency', currency: 'MXN', maximumFractionDigits: 0 }});

            async function initDashboard() {{
                try {{
                    const req = await fetch('https://raw.githubusercontent.com/angelnmara/geojson/master/mexicoHigh.json');
                    geojsonMx = await req.json();
                    updateAll();
                    initDataTable();
                }} catch(e) {{
                    console.error("Error cargando GeoJSON", e);
                }}
            }}

            function getFilteredData() {{
                return rawData.filter(d => {{
                    let pass = true;
                    if(currentFilters.estado && d.Estado !== currentFilters.estado) pass = false;
                    if(currentFilters.ramo && d.Ramo !== currentFilters.ramo) pass = false;
                    if(currentFilters.riesgo && d.Nivel_Riesgo !== currentFilters.riesgo) pass = false;
                    return pass;
                }});
            }}

            function updateAll() {{
                const data = getFilteredData();
                updateKPIs(data);
                updateMap(data);
                updateBar(data);
                updatePie(data);
                updateTop5(data);
                updateTable(data);
                
                // Mostrar u ocultar botón de limpieza
                const btnReset = document.getElementById('btnResetFilters');
                if(currentFilters.estado || currentFilters.ramo || currentFilters.riesgo) {{
                    btnReset.style.display = 'inline-block';
                }} else {{
                    btnReset.style.display = 'none';
                }}
                
                // UI de KPIs activos
                document.getElementById('kpi-card-bajo').classList.toggle('active-filter', currentFilters.riesgo === 'Bajo');
                document.getElementById('kpi-card-medio').classList.toggle('active-filter', currentFilters.riesgo === 'Medio');
                document.getElementById('kpi-card-alto').classList.toggle('active-filter', currentFilters.riesgo === 'Alto');
            }}

            function toggleRiskFilter(riesgoVal) {{
                if(currentFilters.riesgo === riesgoVal) currentFilters.riesgo = null;
                else currentFilters.riesgo = riesgoVal;
                updateAll();
            }}

            function resetAllFilters() {{
                currentFilters.estado = null;
                currentFilters.ramo = null;
                currentFilters.riesgo = null;
                updateAll();
            }}

            function updateKPIs(data) {{
                document.getElementById('kpi-total').innerText = data.length;
                
                const suma = data.reduce((acc, curr) => acc + curr.Monto_Asegurado, 0);
                document.getElementById('kpi-suma').innerText = fmtMoney.format(suma);
                
                document.getElementById('kpi-bajo').innerText = data.filter(d => d.Nivel_Riesgo === 'Bajo').length;
                document.getElementById('kpi-medio').innerText = data.filter(d => d.Nivel_Riesgo === 'Medio').length;
                document.getElementById('kpi-alto').innerText = data.filter(d => d.Nivel_Riesgo === 'Alto').length;
            }}

            function updateMap(data) {{
                if(!geojsonMx) return;
                
                // Agrupar
                let mapData = {{}};
                data.forEach(d => {{
                    let st = d.Estado;
                    // Correcciones mapeo
                    if(st === 'Estado de México') st = 'México';
                    if(st === 'CDMX' || st === 'Distrito Federal') st = 'Ciudad de México';
                    
                    mapData[st] = (mapData[st] || 0) + 1;
                }});

                let locs = Object.keys(mapData);
                let vals = Object.values(mapData);

                let trace = {{
                    type: "choropleth",
                    geojson: geojsonMx,
                    locations: locs,
                    z: vals,
                    featureidkey: "properties.name",
                    colorscale: "Purp",
                    showscale: true,
                    marker: {{ line: {{color: 'white', width: 0.5}} }},
                    colorbar: {{ orientation: "h", x: 0.5, y: -0.1, len: 0.8, thickness: 10, title: "Pólizas" }}
                }};

                let layout = {{
                    geo: {{ fitbounds: "locations", visible: false }},
                    margin: {{l:0, r:0, t:0, b:0}},
                    paper_bgcolor: 'rgba(0,0,0,0)',
                    plot_bgcolor: 'rgba(0,0,0,0)',
                    dragmode: false
                }};

                Plotly.react('plotly-mapa', [trace], layout, {{displayModeBar: false}});
                
                document.getElementById('plotly-mapa').removeAllListeners('plotly_click');
                document.getElementById('plotly-mapa').on('plotly_click', function(data){{
                    let clickedState = data.points[0].location;
                    if(currentFilters.estado === clickedState) currentFilters.estado = null;
                    else currentFilters.estado = clickedState;
                    updateAll();
                }});
            }}

            function updateBar(data) {{
                let grouped = {{}};
                data.forEach(d => {{
                    if(!grouped[d.Ramo]) grouped[d.Ramo] = {{ 'Bajo':0, 'Medio':0, 'Alto':0 }};
                    grouped[d.Ramo][d.Nivel_Riesgo]++;
                }});

                let ramos = Object.keys(grouped);
                let traceBajo = {{ x: ramos, y: ramos.map(r => grouped[r]['Bajo']), name: 'Bajo', type: 'bar', marker: {{color: color_bajo}} }};
                let traceMedio = {{ x: ramos, y: ramos.map(r => grouped[r]['Medio']), name: 'Medio', type: 'bar', marker: {{color: color_medio}} }};
                let traceAlto = {{ x: ramos, y: ramos.map(r => grouped[r]['Alto']), name: 'Alto', type: 'bar', marker: {{color: color_alto}} }};

                let layout = {{
                    barmode: 'stack',
                    margin: {{l:20, r:10, t:20, b:20}},
                    paper_bgcolor: 'rgba(0,0,0,0)',
                    plot_bgcolor: 'rgba(0,0,0,0)',
                    legend: {{orientation: "h", yanchor: "bottom", y: 1.1, xanchor: "center", x: 0.5}},
                    yaxis: {{ showgrid: true, gridcolor: "#E4E7EB" }}
                }};

                Plotly.react('plotly-bar', [traceBajo, traceMedio, traceAlto], layout, {{displayModeBar: false}});
                
                document.getElementById('plotly-bar').removeAllListeners('plotly_click');
                document.getElementById('plotly-bar').on('plotly_click', function(data){{
                    let clickedRamo = data.points[0].x;
                    if(currentFilters.ramo === clickedRamo) currentFilters.ramo = null;
                    else currentFilters.ramo = clickedRamo;
                    updateAll();
                }});
            }}

            function updatePie(data) {{
                let grouped = {{}};
                data.forEach(d => {{ grouped[d.Ramo] = (grouped[d.Ramo] || 0) + 1; }});
                
                let labels = Object.keys(grouped);
                let values = Object.values(grouped);

                let trace = {{
                    type: "pie",
                    labels: labels,
                    values: values,
                    hole: 0.65,
                    marker: {{ colors: ["#311B92", "#5C6BC0", "#9FA8DA", "#D1C4E9"] }}
                }};

                let layout = {{
                    margin: {{l:10, r:10, t:10, b:10}},
                    paper_bgcolor: 'rgba(0,0,0,0)',
                    showlegend: true,
                    legend: {{orientation: "h", y: -0.1}}
                }};

                Plotly.react('plotly-pie', [trace], layout, {{displayModeBar: false}});
            }}

            function updateTop5(data) {{
                // Filtrar solo altos dentro de la data actual
                let altos = data.filter(d => d.Nivel_Riesgo === 'Alto');
                let mapData = {{}};
                altos.forEach(d => {{ mapData[d.Estado] = (mapData[d.Estado] || 0) + 1; }});
                
                let sorted = Object.keys(mapData).map(k => ({{estado: k, count: mapData[k]}})).sort((a,b) => b.count - a.count).slice(0, 5);
                
                let html = '<table class="min-table"><thead><tr><th>Estado</th><th>Pólizas Alto Riesgo</th></tr></thead><tbody>';
                if(sorted.length === 0) html += '<tr><td colspan="2">No hay zonas críticas con los filtros actuales.</td></tr>';
                
                sorted.forEach(s => {{
                    html += `<tr onclick="currentFilters.estado='${{s.estado}}'; updateAll();"><td>${{s.estado}}</td><td>${{s.count}}</td></tr>`;
                }});
                html += '</tbody></table>';
                
                document.getElementById('top5-container').innerHTML = html;
            }}

            function initDataTable() {{
                let cols = [
                    {{ data: "Poliza_ID", title: "Póliza ID" }},
                    {{ data: "Estado", title: "Estado" }},
                    {{ data: "Ramo", title: "Ramo" }},
                    {{ data: "Nivel_Riesgo", title: "Nivel Riesgo" }},
                    {{ 
                        data: "Monto_Asegurado", 
                        title: "Monto Asegurado",
                        render: function(data) {{ return fmtMoney.format(data); }}
                    }}
                ];
                
                dataTableInstance = $('#tablaPolizas').DataTable({{
                    data: rawData,
                    columns: cols,
                    pageLength: 15,
                    responsive: true,
                    language: {{ url: "//cdn.datatables.net/plug-ins/1.13.6/i18n/es-ES.json" }}
                }});
            }}

            function updateTable(data) {{
                if(dataTableInstance) {{
                    dataTableInstance.clear();
                    dataTableInstance.rows.add(data);
                    dataTableInstance.draw();
                }}
            }}

            function openTab(evt, tabName) {{
                let contents = document.getElementsByClassName("tab-content");
                for (let i = 0; i < contents.length; i++) contents[i].style.display = "none";
                
                let links = document.getElementsByClassName("tab-label");
                for (let i = 0; i < links.length; i++) links[i].classList.remove("active");
                
                document.getElementById(tabName).style.display = "block";
                evt.currentTarget.classList.add("active");
                
                window.dispatchEvent(new Event('resize'));
            }}

            window.addEventListener('load', initDashboard);

            window.addEventListener('resize', function() {{
                document.querySelectorAll('.js-plotly-plot').forEach(function(p) {{
                    Plotly.Plots.resize(p);
                }});
            }});
        </script>
    </body>
    </html>"""

    with open("Dashboard_Siniestralidad.html", "w", encoding="utf-8") as f:
        f.write(html_template)
    print("¡Dashboard Generado!")

generar_dashboard_riesgos()
