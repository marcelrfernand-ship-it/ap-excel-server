from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import base64, io, json, traceback
from template_b64 import TEMPLATE_B64

app = Flask(__name__)
CORS(app)

def sc(ws, addr, val):
    """Setar valor em célula preservando formatação original."""
    if val is None or val == '': return
    if addr not in ws: ws[addr] = None
    cell = ws[addr]
    if isinstance(val, (int, float)):
        cell.value = val
    else:
        cell.value = str(val)

@app.route('/', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'AP Excel Generator'})

@app.route('/gerar-excel', methods=['POST'])
def gerar_excel():
    try:
        d = request.get_json(force=True)
        h = d.get('header', {})
        p = d.get('perfil', {})

        # Carregar template
        template_bytes = base64.b64decode(TEMPLATE_B64)
        wb = load_workbook(io.BytesIO(template_bytes))

        # ── PERFIL ──
        ws = wb['🏠 Perfil']
        sc(ws, 'B5', h.get('nome'))
        sc(ws, 'G5', h.get('kam'))
        sc(ws, 'J5', h.get('inside'))
        sc(ws, 'B8', h.get('data'))
        sc(ws, 'E8', h.get('status'))
        sc(ws, 'B13', p.get('receita'))
        sc(ws, 'F13', p.get('assin'))
        sc(ws, 'I13', p.get('cidades'))
        sc(ws, 'L13', p.get('taxa'))
        sc(ws, 'B16', p.get('fat'))
        sc(ws, 'F16', p.get('pos'))
        sc(ws, 'B19', p.get('perspectiva'))
        if p.get('rwdc'): sc(ws, 'B36', float(p['rwdc']))
        if p.get('rpot'): sc(ws, 'C36', float(p['rpot']))
        sc(ws, 'B41', p.get('sativas'))
        sc(ws, 'B46', p.get('salvo'))
        sc(ws, 'B51', p.get('compra'))
        sc(ws, 'B56', p.get('conc'))
        sc(ws, 'B61', p.get('obs'))

        # ── CONTATOS ──
        ws = wb['👥 Contatos']
        for i, ct in enumerate(d.get('contatos', [])[:20]):
            r = 7 + i
            sc(ws, f'B{r}', ct.get('nome'))
            sc(ws, f'C{r}', ct.get('cargo'))
            sc(ws, f'D{r}', ct.get('exec'))
            sc(ws, f'E{r}', ct.get('c'))
            sc(ws, f'F{r}', ct.get('a'))
            sc(ws, f'G{r}', ct.get('inf'))
            sc(ws, f'H{r}', ct.get('tel'))
            sc(ws, f'I{r}', ct.get('email'))
            sc(ws, f'J{r}', ct.get('obs'))

        # ── PROJETOS ──
        ws = wb['📁 Projetos']
        for i, pr in enumerate(d.get('projetos', [])[:20]):
            r = 5 + i
            sc(ws, f'B{r}', pr.get('desc'))
            sc(ws, f'C{r}', pr.get('un'))
            sc(ws, f'D{r}', pr.get('mes'))
            sc(ws, f'E{r}', pr.get('fase'))
            sc(ws, f'F{r}', pr.get('port'))
            if pr.get('valor'): sc(ws, f'G{r}', float(pr['valor']))

        # ── DORES ──
        ws = wb['🎯 Dores']
        for i, dr in enumerate(d.get('dores', [])[:30]):
            r = 7 + i
            sc(ws, f'B{r}', dr.get('dor'))
            sc(ws, f'C{r}', dr.get('un'))
            sc(ws, f'D{r}', dr.get('ini'))
            sc(ws, f'E{r}', dr.get('desc'))
            if dr.get('val'): sc(ws, f'F{r}', float(dr['val']))
            sc(ws, f'G{r}', dr.get('st'))
            sc(ws, f'H{r}', dr.get('cont'))

        # ── AÇÕES ──
        ws = wb['✅ Ações']
        for i, ac in enumerate(d.get('acoes', [])[:25]):
            r = 7 + i
            sc(ws, f'B{r}', ac.get('acao'))
            sc(ws, f'C{r}', ac.get('dor'))
            sc(ws, f'D{r}', ac.get('res'))
            sc(ws, f'E{r}', ac.get('rec'))
            sc(ws, f'F{r}', ac.get('resp'))
            sc(ws, f'G{r}', ac.get('prazo'))
            sc(ws, f'H{r}', ac.get('status'))

        # ── ATAS DE VISITA ──
        ws = wb['📝 Atas de Visita']
        TMAP = {'hot':'🔥 Quente','warm':'🟡 Morno','cold':'❄️ Frio','ana':'🔄 Em análise'}
        for i, vi in enumerate(d.get('visitas', [])[:12]):
            base = 9 + (i * 32)
            sc(ws, f'C{base+1}', vi.get('data'))
            sc(ws, f'D{base+1}', vi.get('hi'))
            sc(ws, f'E{base+1}', vi.get('hf'))
            sc(ws, f'F{base+1}', vi.get('kam'))
            sc(ws, f'H{base+1}', vi.get('tipo'))
            sc(ws, f'C{base+3}', vi.get('objetivo', '').split('\n')[0] if vi.get('objetivo') else '')
            sc(ws, f'F{base+3}', vi.get('kam'))
            sc(ws, f'B{base+5}', vi.get('delta'))
            sc(ws, f'C{base+8}', vi.get('objetivo'))
            sc(ws, f'C{base+11}', vi.get('pontos'))
            sc(ws, f'C{base+15}', vi.get('opps'))
            for j, ac in enumerate(vi.get('acoes', [])[:5]):
                sc(ws, f'C{base+19+j}', ac.get('acao'))
                sc(ws, f'G{base+19+j}', ac.get('resp'))
                sc(ws, f'I{base+19+j}', ac.get('prazo'))
                sc(ws, f'J{base+19+j}', ac.get('status'))
            sc(ws, f'C{base+25}', vi.get('steps'))
            sc(ws, f'G{base+25}', TMAP.get(vi.get('temp',''), ''))
            sc(ws, f'H{base+25}', vi.get('proxContato'))
            sc(ws, f'I{base+25}', '★' * vi.get('score', 0) if vi.get('score') else '')
            sc(ws, f'C{base+28}', vi.get('obsInternas'))

        # ── SEMANAL ──
        if '📊 Semanal' in wb.sheetnames:
            ws = wb['📊 Semanal']
            semanal_data = d.get('semanal', [])
            if isinstance(semanal_data, dict):
                semanal_data = [semanal_data]
            if isinstance(semanal_data, list) and len(semanal_data) > 0:
                last = semanal_data[-1]
                sc(ws, 'B7', last.get('num'))
                sc(ws, 'C7', last.get('periodo'))
                sc(ws, 'E7', h.get('kam'))
                sc(ws, 'B12', last.get('fatos'))
                sc(ws, 'B17', last.get('pipeline'))
                sc(ws, 'F17', last.get('quente'))
                sc(ws, 'B20', last.get('prox'))
                sc(ws, 'F20', last.get('suporte'))
                sc(ws, 'B24', last.get('perdas'))
                sc(ws, 'F24', last.get('conquista'))
                sc(ws, 'B39', last.get('meta'))
                sc(ws, 'D39', last.get('realizado'))
                for i, sem in enumerate(semanal_data[-5:]):
                    row = 30 + i
                    sc(ws, f'B{row}', sem.get('num'))
                    sc(ws, f'E{row}', sem.get('pipeline'))
                    sc(ws, f'G{row}', sem.get('quente'))
                    sc(ws, f'I{row}', sem.get('prox'))

        # ── SEMANAL ──
        if '📊 Semanal' in wb.sheetnames:
            ws = wb['📊 Semanal']
            semanal_data = d.get('semanal', [])
            if isinstance(semanal_data, dict):
                semanal_data = [semanal_data]
            if isinstance(semanal_data, list) and len(semanal_data) > 0:
                last = semanal_data[-1]
                sc(ws, 'B7', last.get('num'))
                sc(ws, 'C7', last.get('periodo'))
                sc(ws, 'E7', h.get('kam'))
                sc(ws, 'B12', last.get('fatos'))
                sc(ws, 'B17', last.get('pipeline'))
                sc(ws, 'F17', last.get('quente'))
                sc(ws, 'B20', last.get('prox'))
                sc(ws, 'F20', last.get('suporte'))
                sc(ws, 'B24', last.get('perdas'))
                sc(ws, 'F24', last.get('conquista'))
                sc(ws, 'B39', last.get('meta'))
                sc(ws, 'D39', last.get('realizado'))
                for i, sem in enumerate(semanal_data[-5:]):
                    row = 30 + i
                    sc(ws, f'B{{row}}', sem.get('num'))
                    sc(ws, f'E{{row}}', sem.get('pipeline'))
                    sc(ws, f'G{{row}}', sem.get('quente'))
                    sc(ws, f'I{{row}}', sem.get('prox'))

                # Salvar em memória e retornar
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        nome = (h.get('nome') or 'conta').replace(' ', '_')
        nome = ''.join(c for c in nome if c.isalnum() or c in '_-')

        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'AP_{nome}.xlsx'
        )

    except Exception as e:
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000)
