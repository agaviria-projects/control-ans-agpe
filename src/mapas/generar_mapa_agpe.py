from pathlib import Path
import pandas as pd
import webbrowser
import json

def generar_mapa_leaflet_agpe():

    # --------------------------------------------------
    # RUTAS
    # --------------------------------------------------
    base_dir = Path(__file__).resolve().parents[2]
    ruta_excel = base_dir / "data_clean" / "AGPE_ANS.xlsm"
    ruta_output = base_dir / "output"
    ruta_output.mkdir(exist_ok=True)
    ruta_html = ruta_output / "mapa_visitas_leaflet.html"

    # --------------------------------------------------
    # LEER EXCEL
    # --------------------------------------------------
    df = pd.read_excel(ruta_excel, dtype=str).fillna("")

    data_js = df[
        [
            "PEDIDO",
            "OBSERVACION",
            "TIPO_VISITA",
            "MUNICIPIO",
            "COORDENADAY",
            "COORDENADAX",
            "DIRECCION",
            "CLIENTE",
            "CELULAR",
        ]
    ].to_dict(orient="records")

    # --------------------------------------------------
    # HTML
    # --------------------------------------------------
    html = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>Mapa Visitas AGPE</title>

<link rel="stylesheet" href="https://unpkg.com/leaflet/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet/dist/leaflet.js"></script>

<style>
html,body{{margin:0;padding:0;height:100%;overflow:hidden;}}
#map{{height:100vh;width:100%;}}

.panelFiltros{{
  position:absolute;
  top:70px;
  right:15px;
  background:#ffffff;
  padding:18px;
  width:270px;
  border-radius:16px;
  box-shadow:0 8px 26px rgba(0,0,0,.25);
  font-family:Segoe UI, Arial;
  z-index:9999;
}}

.panelFiltros h3{{margin:0 0 12px;font-size:16px;color:#1E8449;}}
.lbl{{display:block;margin-top:8px;font-size:12px;font-weight:600;color:#2C3E50;}}

.panelFiltros select,
.inputPedido{{
  width:100%;
  padding:8px;
  margin:6px 0 12px;
  border-radius:8px;
  border:1px solid #ccc;
}}

.btnRow{{display:flex;gap:6px;}}
button{{padding:8px;border:none;border-radius:8px;cursor:pointer;font-size:12px;}}
.btnFiltrar{{background:#3498DB;color:white;}}
.btnLimpiar{{background:#7F8C8D;color:white;}}
.btnCopiar{{background:#2ECC71;color:white;width:100%;margin-top:8px;}}

#toast{{
  position:fixed;
  bottom:30px;
  left:50%;
  transform:translateX(-50%);
  background:#1E8449;
  color:white;
  padding:14px 22px;
  border-radius:12px;
  font-size:13px;
  display:none;
  z-index:99999;
}}
</style>
</head>

<body>

<div id="map"></div>

<div class="panelFiltros">
  <h3>Filtros AGPE</h3>

  <label class="lbl">Pedido</label>
  <input type="text" id="fPedido" class="inputPedido" placeholder="Ej: 23560158">

  <label class="lbl">Observación</label>
  <select id="fObs">
    <option value="">-- Todas --</option>
  </select>

  <label class="lbl">Tipo Visita</label>
  <select id="fTipo">
    <option value="">-- Todas --</option>
  </select>

  <div class="btnRow">
    <button class="btnFiltrar" onclick="aplicarFiltros()">Filtrar</button>
    <button class="btnLimpiar" onclick="limpiarFiltros()">Limpiar</button>
  </div>

  <button class="btnCopiar" onclick="copiarEnlace()">Copiar enlace</button>
</div>

<div id="toast">Datos copiados. Listo para enviar por WhatsApp</div>

<script>
// ======================= MAPA =======================
var map = L.map('map').setView([6.2443,-75.581],12);

var mapaNormal = L.tileLayer(
  'https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png',
  {{
    maxZoom: 18,
    maxNativeZoom: 18
  }}
).addTo(map);

var mapaSatelital = L.tileLayer(
  'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}',
  {{
    maxZoom: 19
  }}
);

L.control.layers(
  {{"Mapa estándar": mapaNormal, "Satelital": mapaSatelital}},
  null,
  {{position:'topright'}}
).addTo(map);

// ======================= ICONOS =======================
var iconAzul = L.icon({{
  iconUrl:'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-blue.png',
  iconSize:[25,41],iconAnchor:[12,41]
}});

var iconNaranja = L.icon({{
  iconUrl:'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-orange.png',
  iconSize:[25,41],iconAnchor:[12,41]
}});

var iconMorado = L.icon({{
  iconUrl:'https://raw.githubusercontent.com/pointhi/leaflet-color-markers/master/img/marker-icon-violet.png',
  iconSize:[25,41],iconAnchor:[12,41]
}});

// ======================= DATOS =======================
var data = {json.dumps(data_js)};
var markers = [];
var layer = L.layerGroup().addTo(map);

// ======================= AGRUPAR POR COORD =======================
var grupos = {{}};
data.forEach(d => {{
  if(d.COORDENADAY && d.COORDENADAX){{
    let k = d.COORDENADAY + ',' + d.COORDENADAX;
    if(!grupos[k]) grupos[k] = [];
    grupos[k].push(d);
  }}
}});

function cargarFiltros(){{
  let obs=new Set(), tipo=new Set();
  data.forEach(d=>{{obs.add(d.OBSERVACION); tipo.add(d.TIPO_VISITA);}});
  obs.forEach(o=>fObs.innerHTML+=`<option value="${{o}}">${{o}}</option>`);
  tipo.forEach(t=>fTipo.innerHTML+=`<option value="${{t}}">${{t}}</option>`);
}}

function pintar(arr){{
  layer.clearLayers(); markers=[];
  let bounds=[];

  Object.keys(grupos).forEach(key => {{
    let pedidos = grupos[key].filter(d => arr.includes(d));
    if(pedidos.length === 0) return;

    let d = pedidos[0];
    let icono = pedidos.length > 1 ? iconMorado : (d.TIPO_VISITA=="C09" ? iconNaranja : iconAzul);

    let popupHtml = "";

    if(pedidos.length > 1){{
      popupHtml = `<b>📍 ${{pedidos.length}} pedidos en esta ubicación</b><br><br>`;
      pedidos.forEach(p => {{
        popupHtml += `
          • <b>Pedido:</b> ${{p.PEDIDO}}<br>
            <b>Cliente:</b> ${{p.CLIENTE}}<br>
            <b>Municipio:</b> ${{p.MUNICIPIO}}<br>
            <b>Tipo visita:</b> ${{p.TIPO_VISITA}}<br>
            <b>Dirección:</b> ${{p.DIRECCION}}<br>
            <b>Coordenadas:</b> ${{p.COORDENADAY}}, ${{p.COORDENADAX}}<br><br>
        `;
      }});
    }} else {{
      popupHtml = `
        <b>Pedido:</b> ${{d.PEDIDO}}<br>
        <b>Cliente:</b> ${{d.CLIENTE}}<br>
        <b>Municipio:</b> ${{d.MUNICIPIO}}<br>
        <b>Tipo visita:</b> ${{d.TIPO_VISITA}}<br>
        <b>Dirección:</b> ${{d.DIRECCION}}<br>
        <b>Celular:</b> ${{d.CELULAR}}<br>
        <b>Coordenadas:</b> ${{d.COORDENADAY}}, ${{d.COORDENADAX}}
      `;
    }}

    let m = L.marker([d.COORDENADAY,d.COORDENADAX],{{icon:icono}})
      .bindPopup(popupHtml);

    m.datos = {{
      pedido: d.PEDIDO,
      cliente: d.CLIENTE,
      municipio: d.MUNICIPIO,
      tipo_visita: d.TIPO_VISITA,
      direccion: d.DIRECCION,
      coordY: d.COORDENADAY,
      coordX: d.COORDENADAX,
      celular: d.CELULAR,
      urlMaps: `https://www.google.com/maps?q=${{d.COORDENADAY}},${{d.COORDENADAX}}`
    }};

    layer.addLayer(m);
    markers.push(m);
    bounds.push([d.COORDENADAY,d.COORDENADAX]);
  }});

  if(bounds.length === 1) {{
    map.setView(bounds[0], 16);
  }} else if(bounds.length > 1) {{
    map.fitBounds(bounds, {{padding:[40,40]}});
  }}
}}

function aplicarFiltros(){{
  pintar(data.filter(d =>
    (!fPedido.value || d.PEDIDO.includes(fPedido.value)) &&
    (!fObs.value || d.OBSERVACION==fObs.value) &&
    (!fTipo.value || d.TIPO_VISITA==fTipo.value)
  ));
}}

function limpiarFiltros(){{
  fPedido.value=""; fObs.value=""; fTipo.value="";
  pintar(data);
}}

// ===================================================
// ✅ COPIAR A PORTAPAPELES (FUNCIONA EN file://)
// ===================================================
function copiarTextoPortapapeles(texto){{
  // 1) Si hay clipboard API y contexto seguro, úsala
  if (navigator.clipboard && window.isSecureContext) {{
    return navigator.clipboard.writeText(texto);
  }}

  // 2) Fallback: textarea + execCommand (sirve en file://)
  return new Promise(function(resolve, reject){{
    try {{
      var ta = document.createElement("textarea");
      ta.value = texto;
      ta.setAttribute("readonly", "");
      ta.style.position = "fixed";
      ta.style.top = "-9999px";
      ta.style.left = "-9999px";
      document.body.appendChild(ta);
      ta.select();
      ta.setSelectionRange(0, ta.value.length);

      var ok = document.execCommand("copy");
      document.body.removeChild(ta);

      if (ok) resolve();
      else reject(new Error("copy_failed"));
    }} catch (err) {{
      reject(err);
    }}
  }});
}}

function copiarEnlace(){{
  let pedido = (fPedido.value || "").trim();
  if(!pedido) {{
    toast.innerText = "Escribe un pedido primero.";
    toast.style.display="block";
    setTimeout(()=>toast.style.display="none",2500);
    return;
  }}

  let m = markers.find(x => x.datos && x.datos.pedido && x.datos.pedido.includes(pedido));
  if(!m) {{
    toast.innerText = "Pedido no encontrado en el mapa actual.";
    toast.style.display="block";
    setTimeout(()=>toast.style.display="none",2500);
    return;
  }}

  let d = m.datos;

  // Texto listo para pegar en WhatsApp
  let texto = `Pedido: ${{d.pedido}}
Cliente: ${{d.cliente}}
Dirección: ${{d.direccion}}
Celular: ${{d.celular}}

Ubicación Google Maps:
${{d.urlMaps}}`;

  copiarTextoPortapapeles(texto)
    .then(function(){{
      toast.innerText = "Datos copiados. Listo para enviar por WhatsApp";
      toast.style.display="block";
      setTimeout(()=>toast.style.display="none",2500);
    }})
    .catch(function(){{
      toast.innerText = "No se pudo copiar. Intenta nuevamente.";
      toast.style.display="block";
      setTimeout(()=>toast.style.display="none",2500);
    }});
}}

cargarFiltros();
pintar(data);
</script>

</body>
</html>
"""

    ruta_html.write_text(html, encoding="utf-8")
    webbrowser.open(ruta_html.as_uri())
