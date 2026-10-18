# CONTAVERSE IA N13 - API REST independiente (Python estándar, sin dependencias)
# Desplegar como SEGUNDO servicio web. No ejecuta ni modifica Streamlit.
import os
import json
import hmac
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

def numero(valor):
    texto = str(valor).strip().replace(" ", "")
    if not texto:
        raise ValueError("Valor numérico vacío")
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        texto = texto.replace(",", ".")
    valor_num = float(texto)
    if not (-float("inf") < valor_num < float("inf")):
        raise ValueError("Valor numérico no finito")
    return valor_num

_A6_ESCENARIOS = (
    ("Base", 1, 0, 0, 0),
    ("Alternativo 1", 2, 15, 1, 0),
    ("Alternativo 2", 3, 40, 2, 1),
    ("Estrés", 4, 80, 3, 1),
)

def _a6_indicadores(datos, dificultad, presion, contradicciones):
    """Modelo reproducible del estado organizacional, sin aleatoriedad."""
    n = len(datos)
    incompletos = sum(not str(d.get("soporte_factura", "")).strip().lower() in ("si", "sí") for d in datos)
    ret_pend = sum(str(d.get("retencion_definida", "")).strip().lower() not in ("si", "sí", "no") for d in datos)
    valores = [numero(d.get("base", 0)) for d in datos]
    if any(x < 0 for x in valores):
        raise ValueError("Las bases de operación deben ser no negativas.")
    facturacion = round(sum(valores), 2)
    complejidad = dificultad * n + incompletos * 2 + ret_pend + contradicciones * n
    carga = complejidad * (1 + presion / 100)
    capacidad = max(1, n * 3)
    saturacion = round(100 * carga / capacidad, 2)
    riesgo = min(100, round(8 + 9*dificultad + 0.25*presion + 8*contradicciones + (35*incompletos/max(n, 1)) + (15*ret_pend/max(n, 1)), 2))
    tiempo_estimado = round(carga * 4, 2)
    casos_escalados = min(n, int((n * riesgo / 100) + 0.9999))
    if riesgo >= 75 or saturacion >= 120:
        dictamen = "ESCALAR Y APLAZAR decisión hasta revisión de evidencias"
    elif riesgo >= 45 or saturacion >= 90:
        dictamen = "MODIFICAR condiciones y solicitar revisión humana"
    else:
        dictamen = "REVISIÓN HUMANA previa a cualquier aprobación"
    return {
        "Casos": n, "Valor base total": facturacion,
        "Soportes pendientes": incompletos, "Retenciones sin definir": ret_pend,
        "Dificultad": dificultad, "Presión de tiempo (%)": presion,
        "Contradicciones por caso": contradicciones,
        "Carga estimada (unidades)": round(carga, 2),
        "Capacidad de referencia (unidades)": capacidad,
        "Saturación (%)": saturacion,
        "Riesgo modelado (0-100)": riesgo,
        "Tiempo estimado de revisión (min)": tiempo_estimado,
        "Casos sugeridos para escalamiento": casos_escalados,
        "Recomendación automatizada": dictamen,
    }

EJEMPLO = [
    {"descripcion":"Compra de inventario","base":"12000000","soporte_factura":"si","retencion_definida":"si"},
    {"descripcion":"Servicio sin factura","base":"7200000","soporte_factura":"no","retencion_definida":""},
    {"descripcion":"Compra de maquinaria","base":"38000000","soporte_factura":"si","retencion_definida":"no"},
    {"descripcion":"Consultoría","base":"9500000","soporte_factura":"si","retencion_definida":"si"},
]

def calcular(datos, dificultad=1, presion=0, contradicciones=0):
    if not isinstance(datos,list) or not 1 <= len(datos) <= 5000 or not all(isinstance(f,dict) for f in datos):
        raise ValueError("Se necesita una lista de 1 a 5000 casos (objetos JSON).")
    if not (1 <= dificultad <= 4 and 0 <= presion <= 100 and 0 <= contradicciones <= 3):
        raise ValueError("Parámetros fuera de rango.")
    gemelo = _a6_indicadores(datos, dificultad, presion, contradicciones)
    escenarios = []
    for nombre, d, p, c, _ in _A6_ESCENARIOS:
        r = _a6_indicadores(datos, d, p, c)
        escenarios.append({"Escenario":nombre,"Dificultad":d,"Presión (%)":p,"Contradicciones":c,
            "Riesgo (0-100)":r["Riesgo modelado (0-100)"],"Saturación (%)":r["Saturación (%)"],
            "Revisión (min)":r["Tiempo estimado de revisión (min)"],
            "Casos escalados":r["Casos sugeridos para escalamiento"],
            "Recomendación motor":r["Recomendación automatizada"]})
    return {"n11":gemelo,"n12":escenarios,"origen":"Reglas locales determinísticas",
        "advertencia":"Índices educativos, no dictamen profesional ni IA generativa."}

class Handler(BaseHTTPRequestHandler):
    def enviar(self, status, obj):
        datos=json.dumps(obj,ensure_ascii=False,allow_nan=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type','application/json; charset=utf-8')
        self.send_header('Content-Length',str(len(datos)))
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Headers','Content-Type, X-API-Key')
        self.send_header('Access-Control-Allow-Methods','GET, POST, OPTIONS')
        self.send_header('Cache-Control','no-store')
        self.end_headers(); self.wfile.write(datos)

    def autenticado(self):
        clave=os.environ.get('CONTAVERSE_API_KEY','')
        return not clave or hmac.compare_digest(self.headers.get('X-API-Key',''),clave)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header('Access-Control-Allow-Origin','*')
        self.send_header('Access-Control-Allow-Headers','Content-Type, X-API-Key')
        self.send_header('Access-Control-Allow-Methods','GET, POST, OPTIONS')
        self.end_headers()

    def do_GET(self):
        ruta=urlparse(self.path).path.rstrip('/') or '/'
        if ruta=='/health':
            return self.enviar(200,{"estado":"ok","servicio":"CONTAVERSE N13 API","datos":"ninguno"})
        if not self.autenticado(): return self.enviar(401,{"error":"Clave API no válida"})
        try:
            if ruta in ('/api/n11','/api/n12','/api/resumen'):
                res=calcular(EJEMPLO)
                salida=res['n11'] if ruta=='/api/n11' else res['n12'] if ruta=='/api/n12' else res
                return self.enviar(200,{"tipo":"DEMO_SINTETICA","fuente":"Datos de ejemplo del archivo original",
                                        "resultado":salida,"nota":"No son datos en vivo de la sesión de Streamlit."})
            return self.enviar(404,{"error":"Ruta no encontrada"})
        except (ValueError,TypeError,KeyError) as e:
            return self.enviar(400,{"error":str(e)})

    def do_POST(self):
        ruta=urlparse(self.path).path.rstrip('/')
        if ruta!='/api/simular': return self.enviar(404,{"error":"Ruta no encontrada"})
        if not self.autenticado(): return self.enviar(401,{"error":"Clave API no válida"})
        try:
            longitud=int(self.headers.get('Content-Length','0'))
            if longitud<1 or longitud>2_000_000:
                return self.enviar(413,{"error":"Cuerpo ausente o mayor a 2 MB"})
            payload=json.loads(self.rfile.read(longitud))
            if not isinstance(payload,dict): raise ValueError('El JSON debe ser un objeto')
            datos=payload.get('datos')
            dificultad=payload.get('dificultad',1)
            presion=payload.get('presion',0)
            contradicciones=payload.get('contradicciones',0)
            if not all(type(x) is int for x in (dificultad,presion,contradicciones)):
                raise ValueError('Los controles deben ser números enteros')
            return self.enviar(200,{"tipo":"SOLICITUD_REAL_API","resultado":calcular(datos,dificultad,presion,contradicciones)})
        except (ValueError,TypeError,KeyError,json.JSONDecodeError) as e:
            return self.enviar(400,{"error":str(e)})

if __name__=='__main__':
    puerto=int(os.environ.get('PORT','8000'))
    print(f'CONTAVERSE N13 API en puerto {puerto}',flush=True)
    ThreadingHTTPServer(('0.0.0.0',puerto),Handler).serve_forever()
