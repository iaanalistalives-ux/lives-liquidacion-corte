# -*- coding: utf-8 -*-
"""Lee el reporte "RESUMEN LIQUIDACION CORTE - ORDENES DE CORTE TERMINADAS" (Excel del sistema)
y lo deja listo para la base (accion "cargar" de liqc_api).

Uso (lo corre Claude en la PC; el Excel nunca va a GitHub):
  python -I cargar_liquidacion.py "<ruta del .xlsx>"                 -> solo revisa y muestra el resumen
  python -I cargar_liquidacion.py "<ruta del .xlsx>" --json sal.json -> ademas guarda el JSON
  python -I cargar_liquidacion.py "<ruta del .xlsx>" --enviar        -> lo carga en la base
  python -I cargar_liquidacion.py "<ruta del .xlsx>" --enviar --sin-alertas
                                   -> lo carga sin generar alertas (para un reporte historico)

Al cargar, la base compara cada tela con lo que habia y anota las alertas para Alertas LIVES
(tela liquidada, volvio a por liquidar, cambio de porcentaje). La respuesta dice cuantas salieron.

El calculo del cuadro (celdas W44, W45 y W46) NO se lee del Excel: lo hace la base con cada linea.
  total       = lineas con dato en "% DESBALANCE" de CORTE PIEZA
  liquidadas  = lineas con ese % menor a 1 (100%)
  % liquidado = liquidadas / total
"""
import sys, json, datetime, urllib.request, urllib.error
import openpyxl

SB = 'https://xxfufpddgoyhbotjmaiy.supabase.co'
KEY = 'sb_publishable_1g-n6UT70EUlNAo8Lyc04w_jKR4LpXZ'
USUARIO = 'ctalepcio'

# columna del Excel (letra) -> campo de la base
TEXTO = {'B': 'grupo', 'C': 'estilo', 'D': 'oc', 'E': 'op', 'X': 'modulo_liq', 'Z': 'partida'}
NUMERO = {'I': 'pa_ingreso', 'J': 'pa_dev_almacen', 'K': 'pa_dev_corte', 'L': 'pa_tendido', 'M': 'pa_retazos',
          'N': 'pa_puntas', 'O': 'pa_desbalance', 'P': 'pz_ingreso', 'Q': 'pz_dev_almacen', 'R': 'pz_dev_corte',
          'S': 'pz_tendido', 'T': 'pz_retazos', 'U': 'pz_puntas', 'V': 'pz_empalmes', 'W': 'pz_desbalance'}
ENTERO = {'AA': 'giradas', 'AB': 'cortadas', 'AC': 'habilitadas', 'AD': 'merma_corte'}
# encabezados que deben estar donde se espera (fila de titulos y la siguiente); si el reporte cambia, se detiene
ESPERADO = {('B', 0): 'GRUPO', ('D', 0): 'OC', ('E', 0): 'OP', ('F', 0): 'TELA', ('G', 0): 'COLOR', ('H', 0): 'FECHA',
            ('I', 1): 'INGRESO', ('L', 1): 'TENDIDO', ('O', 1): '% DESBALANCE', ('P', 1): 'INGRESO',
            ('V', 1): 'EMPALMES', ('W', 1): '% DESBALANCE', ('Y', 0): 'MODULO', ('Z', 0): 'PARTIDA',
            ('AA', 1): 'GIRADAS', ('AB', 1): 'CORTADAS', ('AC', 1): 'HABILITADAS'}


def txt(v):
    return '' if v is None else str(v).strip()


def numero(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(',', '')
    if s == '':
        return None
    return float(s)


def partir(v):
    """'JE002165-Jersey ...' -> ('JE002165', 'Jersey ...'). Sin guion: codigo vacio."""
    s = txt(v)
    if '-' not in s:
        return '', s
    c, n = s.split('-', 1)
    return c.strip().upper(), n.strip()


def a_fecha(v):
    if isinstance(v, datetime.datetime):
        return v.date().isoformat()
    if isinstance(v, datetime.date):
        return v.isoformat()
    s = txt(v)
    for f in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
        try:
            return datetime.datetime.strptime(s, f).date().isoformat()
        except ValueError:
            pass
    raise ValueError('Fecha que no se entiende: %r' % (v,))


def leer(ruta):
    ws = openpyxl.load_workbook(ruta, data_only=True).worksheets[0]
    cel = lambda col, fila: ws['%s%d' % (col, fila)].value
    cab = next((f for f in range(1, 30) if txt(cel('B', f)).upper() == 'GRUPO'), None)
    if not cab:
        raise SystemExit('No encuentro la fila de titulos (GRUPO en la columna B). No es el reporte esperado.')
    for (col, d), nombre in ESPERADO.items():
        if txt(cel(col, cab + d)).upper() != nombre:
            raise SystemExit('El reporte cambio de forma: en %s%d esperaba "%s" y hay "%s". No se carga nada.'
                             % (col, cab + d, nombre, txt(cel(col, cab + d))))
    desde = hasta = None
    for f in range(1, cab):
        for c in range(1, 12):
            v = txt(ws.cell(f, c).value).upper()
            if v.startswith('DESDE'):
                desde = a_fecha(ws.cell(f, c + 1).value)
            if v.startswith('HASTA'):
                hasta = a_fecha(ws.cell(f, c + 1).value)
    lineas, ocultas = [], 0
    for f in range(cab + 2, ws.max_row + 1):
        if not txt(cel('D', f)):
            continue                       # filas vacias y las de calculos propios al pie (sin OC). El grupo en blanco SI entra
        if ws.row_dimensions[f].hidden:    # regla: solo filas visibles
            ocultas += 1
            continue
        r = {'orden': len(lineas) + 1}
        for col, campo in TEXTO.items():
            r[campo] = txt(cel(col, f))
        r['grupo'] = r['grupo'].upper()
        r['tela_cod'], r['tela'] = partir(cel('F', f))
        r['color_cod'], r['color'] = partir(cel('G', f))
        r['modulo_cod'], r['modulo'] = partir(cel('Y', f))
        r['fecha'] = a_fecha(cel('H', f))
        for col, campo in NUMERO.items():
            r[campo] = numero(cel(col, f))
        for col, campo in ENTERO.items():
            v = numero(cel(col, f))
            r[campo] = None if v is None else int(round(v))
        lineas.append(r)
    if not lineas:
        raise SystemExit('El reporte no trae lineas.')
    lineas, repetidas = sin_repetidas(lineas)
    fechas = [r['fecha'] for r in lineas]
    desde, hasta = desde or min(fechas), hasta or max(fechas)
    return {'desde': desde, 'hasta': hasta, 'lineas': lineas}, ocultas, repetidas


def sin_repetidas(lineas):
    """El reporte a veces trae la misma tela de la misma orden de corte dos veces (misma OC, tela, color y
    partida, con las mismas cantidades; solo cambia el modulo de liquidacion, una en blanco). Entra una sola
    vez: la que trae el modulo de liquidacion. Si las cantidades NO son iguales, no se toca: entran las dos."""
    fuera = ('orden', 'modulo_liq')
    vistas, quitar, repetidas = {}, set(), []
    for i, r in enumerate(lineas):
        clave = (r['oc'], r['tela_cod'].upper(), r['color_cod'].upper(), r['partida'])
        igual = tuple((k, v) for k, v in sorted(r.items()) if k not in fuera)
        j = vistas.get((clave, igual))
        if j is None:
            vistas[(clave, igual)] = i
            continue
        if r['modulo_liq'] and not lineas[j]['modulo_liq']:
            quitar.add(j); vistas[(clave, igual)] = i
        else:
            quitar.add(i)
        repetidas.append('OC %s - tela %s - partida %s' % (r['oc'], r['tela_cod'], r['partida']))
    queda = [r for i, r in enumerate(lineas) if i not in quitar]
    for n, r in enumerate(queda, 1):
        r['orden'] = n
    return queda, repetidas


def resumen(d):
    cuentan = [r for r in d['lineas'] if r['pz_desbalance'] is not None]
    liq = [r for r in cuentan if r['pz_desbalance'] < 1]
    return {'lineas': len(d['lineas']), 'total': len(cuentan), 'liquidadas': len(liq),
            'por_liquidar': len(cuentan) - len(liq), 'pct': (len(liq) / len(cuentan)) if cuentan else None,
            'ordenes_de_corte': len({r['oc'] for r in d['lineas']})}


def enviar(d, archivo, sin_alertas=False):
    p = dict(d, usuario=USUARIO, archivo=archivo)
    if sin_alertas:
        p['sin_alertas'] = 'si'
    cuerpo = json.dumps({'p_accion': 'cargar', 'p': p}).encode('utf-8')
    req = urllib.request.Request(SB + '/rest/v1/rpc/liqc_api', data=cuerpo, method='POST',
                                 headers={'Content-Type': 'application/json', 'apikey': KEY})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        raise SystemExit('La base rechazo la carga (%s): %s' % (e.code, e.read().decode('utf-8', 'replace')))


def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    ruta = sys.argv[1]
    d, ocultas, repetidas = leer(ruta)
    s = resumen(d)
    print('Reporte del %s al %s' % (d['desde'], d['hasta']))
    print('Lineas: %(lineas)d (%(ordenes_de_corte)d ordenes de corte) | cuentan: %(total)d | liquidadas: %(liquidadas)d | por liquidar: %(por_liquidar)d' % s)
    print('%% liquidado: %s' % ('sin dato' if s['pct'] is None else '%.2f%%' % (s['pct'] * 100)))
    if ocultas:
        print('AVISO: %d filas ocultas con datos no se cargaron (regla: solo filas visibles).' % ocultas)
    if repetidas:
        print('AVISO: %d fila(s) repetida(s) del reporte entraron una sola vez: %s' % (len(repetidas), '; '.join(repetidas)))
    if '--json' in sys.argv:
        sal = sys.argv[sys.argv.index('--json') + 1]
        with open(sal, 'w', encoding='utf-8') as f:
            json.dump(d, f, ensure_ascii=False)
        print('JSON guardado en', sal)
    if '--enviar' in sys.argv:
        nombre = ruta.replace('\\', '/').split('/')[-1]
        r = enviar(d, nombre, '--sin-alertas' in sys.argv)
        print('Respuesta de la base:', json.dumps(r, ensure_ascii=False))
        if r.get('cambios_error'):
            print('AVISO: la carga se hizo, pero la comparacion para Alertas LIVES fallo:', r['cambios_error'])
        elif '--sin-alertas' in sys.argv:
            print('Alertas LIVES: carga sin alertas (a pedido).')
        else:
            print('Alertas LIVES: %s alerta(s) de liquidacion.' % r.get('cambios', 'sin dato (la base aun no compara)'))


if __name__ == '__main__':
    main()
