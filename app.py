from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import json, io, base64, shutil, os, tempfile
from openpyxl import load_workbook

app = Flask(__name__)
CORS(app)  # Allow requests from Netlify

# Template is embedded as base64 in the server
TEMPLATE_B64 = open('/app/template_b64.txt').read().strip()

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'service': 'KN Papers Indent API'})

@app.route('/generate-indent', methods=['POST'])
def generate_indent():
    try:
        data = request.json
        
        # Decode template
        template_bytes = base64.b64decode(TEMPLATE_B64)
        template_io = io.BytesIO(template_bytes)
        
        # Load workbook
        wb = load_workbook(template_io)
        ws = wb.active
        
        def W(cell, value):
            try: ws[cell] = value
            except: pass
        
        # Extract data
        iNo    = data.get('iNo', '')
        iMonth = data.get('iMonth', 'MAY').upper()
        iDate  = data.get('iDate', '')
        d      = data.get('dealer', {})
        c      = data.get('consignee') or d
        b      = data.get('buyer') or d
        items  = data.get('items', [])
        trans  = data.get('trans', 'ROAD')
        ins    = data.get('ins', 'BY CONSIGNEE')
        deliv  = data.get('deliv', 'IMMEDIATE/ STANDARD')
        
        # Fill cells - exact positions from original XLS
        W('C4', iMonth); W('D4', iNo); W('C5', iDate)
        
        # Dealer
        W('C7', d.get('name','')); W('C8', d.get('addr',''))
        W('C9', 'Ph No: ' + d.get('ph',''))
        W('B11', d.get('gstin','')); W('F11', d.get('st',''))
        
        # Consignee
        W('D13', c.get('name',d.get('name','')))
        W('C14', c.get('addr',d.get('addr','')))
        W('C15', 'Ph No: ' + c.get('ph',d.get('ph','')))
        W('B18', c.get('gstin',d.get('gstin',''))); W('F18', c.get('st',d.get('st','')))
        
        # Buyer
        W('C20', b.get('name',d.get('name','')))
        W('C21', b.get('addr',d.get('addr','')))
        W('C22', 'Ph No: ' + b.get('ph',d.get('ph','')))
        W('B24', b.get('gstin',d.get('gstin',''))); W('F24', b.get('st',d.get('st','')))
        
        # Transport & Insurance
        W('L23', trans)
        if ins == 'SELF':
            W('D26', 'SELF  ✔'); W('J26', '')
        else:
            W('D26', 'SELF'); W('J26', '✔')
        
        # Items rows 30-37
        for i, it in enumerate(items[:8]):
            r = 30 + i
            W(f'A{r}', it.get('sl', i+1))
            W(f'B{r}', it.get('mc', ''))
            W(f'C{r}', it.get('code', ''))
            W(f'D{r}', it.get('prod', ''))
            W(f'F{r}', it.get('hsn', ''))
            W(f'G{r}', it.get('gsm', ''))
            W(f'H{r}', it.get('sz', ''))
            try: W(f'I{r}', float(it.get('rw','')) if it.get('rw') else '')
            except: W(f'I{r}', it.get('rw',''))
            W(f'J{r}', it.get('ns', ''))
            W(f'K{r}', it.get('col', ''))
            try: W(f'L{r}', float(it.get('qty','')) if it.get('qty') else '')
            except: W(f'L{r}', it.get('qty',''))
            W(f'M{r}', it.get('deck', ''))
            W(f'N{r}', it.get('rem', ''))
        
        # Fill remaining empty rows
        for i in range(len(items), 8):
            r = 30 + i
            for col in ['A','B','C','D','F','G','H','I','J','K','L','M','N']:
                W(f'{col}{r}', '')
        
        # Special instructions
        W('D41', deliv)
        W('L41', 'ALLOW / DISALLOW')
        W('L42', 'ALLOW / DISALLOW/NOT APPLICABLE')
        W('L43', 'ALLOW / DISALLOW/NOT APPLICABLE')
        W('L44', 'ALLOW / DISALLOW/NOT APPLICABLE')
        W('L45', 'ALLOW / DISALLOW')
        W('K49', 'For: ' + d.get('name', 'SHRI K.N.PAPERS'))
        
        # Save to buffer
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Return as downloadable file
        filename = f"WCPM_Indent_{iMonth}_{iNo}.xlsx"
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
