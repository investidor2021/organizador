// script.js (versão melhorada: extenso correto + DOCX/PDF formatados)
// Depende de: XLSX, jspdf, docx já carregados no index.html

// -----------------------------
// UTILITÁRIOS: formatação / parse
// -----------------------------
function parseNumberFromString(str) {
    if (str === null || str === undefined || str === '') return null;
    let s = String(str).replace(/\s/g, '').replace(/\u00A0/g, '').replace(/\./g, '').replace(',', '.');
    s = s.replace(/[^\d\.\-]/g, '');
    const n = Number(s);
    return isNaN(n) ? null : n;
}
function formatCurrency(value) {
    if (value === null || value === undefined || value === '') return '';
    const n = Number(value);
    if (isNaN(n)) return '';
    return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// -----------------------------
// EXTENSO: implementação robusta (pt-BR) para reais e centavos
// -----------------------------
function numeroParaExtensoBR(valor) {
    // aceita number ou string (com vírgula ou ponto)
    if (valor === null || valor === undefined || valor === '') return '';
    const n = parseNumberFromString(valor);
    if (n === null) return '';

    const inteiro = Math.floor(Math.abs(n));
    const centavos = Math.round((Math.abs(n) - inteiro) * 100);

    function leitura3(n) {
        const U = ['','um','dois','três','quatro','cinco','seis','sete','oito','nove','dez','onze','doze','treze','quatorze','quinze','dezesseis','dezessete','dezoito','dezenove'];
        const D = ['','','vinte','trinta','quarenta','cinquenta','sessenta','setenta','oitenta','noventa'];
        const C = ['','cento','duzentos','trezentos','quatrocentos','quinhentos','seiscentos','setecentos','oitocentos','novecentos'];

        if (n === 0) return '';
        if (n === 100) return 'cem';
        let texto = '';
        const c = Math.floor(n / 100);
        const r = n % 100;
        if (c) texto += C[c];
        if (r) {
            if (texto) texto += ' e ';
            if (r < 20) texto += U[r];
            else {
                const d = Math.floor(r / 10);
                const u = r % 10;
                texto += D[d];
                if (u) texto += ' e ' + U[u];
            }
        }
        return texto;
    }

    function juntarPartes(partes) {
        // partes já em ordem natural (milhões...centenas)
        if (partes.length === 0) return '';
        if (partes.length === 1) return partes[0];
        // inserir vírgula entre grupos maiores (se houver) e ' e ' antes do último quando apropriado
        if (partes.length === 2) return partes[0] + ' e ' + partes[1];
        // 3+ partes -> separar com ', ' e ' e ' antes da última
        const allButLast = partes.slice(0, -1).join(', ');
        return allButLast + ' e ' + partes[partes.length - 1];
    }

    if (inteiro === 0) {
        var textoInteiro = 'zero reais';
    } else {
        // quebra em grupos de 3 dígitos
        const sufSing = ['', 'mil', 'milhão', 'bilhão', 'trilhão'];
        const sufPlur = ['', 'mil', 'milhões', 'bilhões', 'trilhões'];

        let x = inteiro;
        const grupos = [];
        while (x > 0) {
            grupos.push(x % 1000);
            x = Math.floor(x / 1000);
        }
        const partes = [];
        for (let i = grupos.length - 1; i >= 0; i--) {
            const g = grupos[i];
            if (g === 0) continue;
            let txt = leitura3(g);
            if (i > 0) {
                // tratamento para 'um mil' -> apenas 'mil'
                if (g === 1 && i === 1) txt = 'mil';
                else txt += (g === 1 ? ' ' + sufSing[i] : ' ' + sufPlur[i]);
            }
            partes.push(txt);
        }
        textoInteiro = juntarPartes(partes) + (Math.abs(Math.floor(n)) === 1 ? ' real' : ' reais');
    }

    let resultado = textoInteiro;
    if (centavos && centavos > 0) {
        const centTxt = (function(c) {
            if (c === 0) return '';
            if (c < 100) return leitura3(c);
            return leitura3(c); // c < 100 anyway
        })(centavos);
        resultado += ' e ' + centTxt + (centavos === 1 ? ' centavo' : ' centavos');
    }

    if (n < 0) resultado = 'menos ' + resultado;
    return resultado;
}

// -----------------------------
// EVENT: formata input moeda enquanto digita
// -----------------------------
function attachCurrencyFormatter(input) {
    input.addEventListener('input', (e) => {
        const el = e.target;
        const digits = el.value.replace(/\D/g, '');
        if (!digits) { el.value = ''; return; }
        const num = parseInt(digits, 10) / 100;
        el.value = num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    });
}

// -----------------------------
// LÓGICA PRINCIPAL: leitura Excel + UI
// -----------------------------
document.addEventListener('DOMContentLoaded', () => {
    // elementos esperados no index.html
    const excelFileInput = document.getElementById('excel-file');
    const fichasCount = document.getElementById('fichas-carregadas-count');
    const excelDisplay = document.getElementById('excel-data-display');

    const superavitCheck = document.getElementById('fonte_superavit');
    const excessoCheck = document.getElementById('fonte_excesso');
    const anulacaoCheck = document.getElementById('fonte_anulacao');

    const superavitSection = document.getElementById('superavit-valor-section');
    const excessoSection = document.getElementById('excesso-valor-section');
    const anulacaoSection = document.getElementById('anulacao-section');
    const creditoSection = document.getElementById('credito-section');

    const addAnulacaoBtn = document.getElementById('add-anulacao-ficha-btn');
    const addCreditoBtn = document.getElementById('add-credito-ficha-btn');
    const anulacaoContainer = document.getElementById('anulacao-fichas-container');
    const creditoContainer = document.getElementById('credito-fichas-container');

    const processarBtn = document.getElementById('processar-btn');
    const gerarPdfBtn = document.getElementById('gerar-pdf-btn');
    const gerarDocxBtn = document.getElementById('gerar-docx-btn');
    const projetoLeiContainer = document.getElementById('projeto-lei-gerado');

    const valorSuperavitInput = document.getElementById('valor-superavit');
    const valorExcessoInput = document.getElementById('valor-excesso');

    // assinatura opcional
    let nomeSecretariaInput = document.getElementById('nome-secretaria');
    let cargoSecretariaInput = document.getElementById('cargo-secretaria');

    // array de fichas lidas
    let fichasExcel = []; // { codigo, descricao, valor }

    // atualiza visibilidade
    function atualizarVisibilidade() {
        superavitSection?.classList.toggle('hidden', !superavitCheck?.checked);
        excessoSection?.classList.toggle('hidden', !excessoCheck?.checked);
        anulacaoSection?.classList.toggle('hidden', !anulacaoCheck?.checked);
        const anyFonte = superavitCheck?.checked || excessoCheck?.checked || anulacaoCheck?.checked;
        creditoSection?.classList.toggle('hidden', !anyFonte);
        addAnulacaoBtn.disabled = !anulacaoCheck.checked || fichasExcel.length === 0;
        addCreditoBtn.disabled = !anyFonte || fichasExcel.length === 0;
    }
    [superavitCheck, excessoCheck, anulacaoCheck].forEach(chk => { if (chk) chk.addEventListener('change', atualizarVisibilidade); });

    // anexa formatadores iniciais de inputs de valor (se existirem)
    if (valorSuperavitInput) attachCurrencyFormatter(valorSuperavitInput);
    if (valorExcessoInput) attachCurrencyFormatter(valorExcessoInput);

    // leitura do Excel (código completo na coluna 0)
    excelFileInput?.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            const sheetName = wb.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

            // detectar cabeçalho
            let start = 0;
            if (rows.length && rows[0].some(c => /cod|ficha|cód|código|codigo|descr/i.test(String(c).toLowerCase()))) start = 1;

            fichasExcel = [];
            for (let i = start; i < rows.length; i++) {
                const r = rows[i];
                if (!r || r.length === 0) continue;
                const codigo = r[0] !== undefined ? String(r[0]).trim() : '';
                const descricao = r[1] !== undefined ? String(r[1]).trim() : '';
                let valor = null;
                for (let c = 2; c < r.length; c++) {
                    const maybe = parseNumberFromString(r[c]);
                    if (maybe !== null) { valor = maybe; break; }
                }
                if (valor === null && r[2] !== undefined) valor = parseNumberFromString(r[2]);
                if (codigo || descricao) fichasExcel.push({ codigo: codigo || descricao, descricao: descricao || '', valor: valor });
            }

            fichasCount.textContent = `Fichas carregadas: ${fichasExcel.length}`;
            excelDisplay?.classList.remove('hidden');
            atualizarVisibilidade();
        };
        reader.readAsArrayBuffer(file);
    });

    function optionTextFromFicha(f) {
        const v = (f.valor !== null && f.valor !== undefined) ? ` - R$ ${formatCurrency(f.valor)}` : '';
        return `${f.codigo}${f.descricao ? ' - ' + f.descricao : ''}${v}`;
    }

    // adicionar anulação
    addAnulacaoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const row = document.createElement('div'); row.className = 'anul-item'; row.style.marginBottom = '8px';
        const sel = document.createElement('select'); sel.style.minWidth = '420px';
        fichasExcel.forEach((f, idx) => { const o=document.createElement('option'); o.value = idx; o.textContent = optionTextFromFicha(f); sel.appendChild(o); });
        const valor = document.createElement('input'); valor.type='text'; valor.placeholder='Valor (R$)'; valor.size=12; valor.style.marginLeft='8px';
        attachCurrencyFormatter(valor);
        const rem = document.createElement('button'); rem.type='button'; rem.textContent='Remover'; rem.style.marginLeft='8px';
        rem.addEventListener('click', ()=>{ row.remove(); calcularTotais(); });
        sel.addEventListener('change', ()=>{ const f = fichasExcel[Number(sel.value)]; if (f && f.valor!==null) valor.value = formatCurrency(f.valor); calcularTotais(); });
        if (fichasExcel[0] && fichasExcel[0].valor!==null) valor.value = formatCurrency(fichasExcel[0].valor);
        row.appendChild(document.createTextNode('Origem: ')); row.appendChild(sel); row.appendChild(document.createTextNode(' Valor: ')); row.appendChild(valor); row.appendChild(rem);
        anulacaoContainer.appendChild(row); calcularTotais();
    });

    // adicionar credito
    addCreditoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const row = document.createElement('div'); row.className='cred-item'; row.style.marginBottom='8px';
        const sel = document.createElement('select'); sel.style.minWidth='420px';
        fichasExcel.forEach((f, idx) => { const o=document.createElement('option'); o.value=idx; o.textContent=optionTextFromFicha(f); sel.appendChild(o); });
        const valor = document.createElement('input'); valor.type='text'; valor.placeholder='Valor (R$)'; valor.size=12; valor.style.marginLeft='8px';
        attachCurrencyFormatter(valor);
        const rem = document.createElement('button'); rem.type='button'; rem.textContent='Remover'; rem.style.marginLeft='8px';
        rem.addEventListener('click', ()=>{ row.remove(); calcularTotais(); });
        sel.addEventListener('change', ()=>{ const f = fichasExcel[Number(sel.value)]; if (f && f.valor!==null) valor.value = formatCurrency(f.valor); calcularTotais(); });
        if (fichasExcel[0] && fichasExcel[0].valor!==null) valor.value = formatCurrency(fichasExcel[0].valor);
        row.appendChild(document.createTextNode('Destino: ')); row.appendChild(sel); row.appendChild(document.createTextNode(' Valor: ')); row.appendChild(valor); row.appendChild(rem);
        creditoContainer.appendChild(row); calcularTotais();
    });

    // calcular totais
    function calcularTotais() {
        let totalAnul = 0; let totalCred = 0;
        anulacaoContainer.querySelectorAll('input').forEach(inp => { const v = parseNumberFromString(inp.value); if (v !== null) totalAnul += v; });
        creditoContainer.querySelectorAll('input').forEach(inp => { const v = parseNumberFromString(inp.value); if (v !== null) totalCred += v; });
        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        let totAnEl = document.getElementById('total-anulacao-display'); let totCrEl = document.getElementById('total-credito-display');
        if (!totAnEl) { totAnEl = document.createElement('div'); totAnEl.id='total-anulacao-display'; anulacaoContainer.parentNode.insertBefore(totAnEl, anulacaoContainer.nextSibling); }
        if (!totCrEl) { totCrEl = document.createElement('div'); totCrEl.id='total-credito-display'; creditoContainer.parentNode.insertBefore(totCrEl, creditoContainer.nextSibling); }
        totAnEl.innerHTML = `<strong>Total Anulação:</strong> R$ ${formatCurrency(totalAnul.toFixed(2))} (${numeroParaExtensoBR(totalAnul.toFixed(2))})`;
        const totalFontes = totalAnul + superVal + excessoVal;
        totCrEl.innerHTML = `<strong>Total Crédito:</strong> R$ ${formatCurrency(totalCred.toFixed(2))} (${numeroParaExtensoBR(totalCred.toFixed(2))}) <br><strong>Total Fontes:</strong> R$ ${formatCurrency(totalFontes.toFixed(2))} (${numeroParaExtensoBR(totalFontes.toFixed(2))})`;
    }

    document.addEventListener('input', (e) => {
        if (e.target && e.target.tagName === 'INPUT' && e.target.type === 'text') calcularTotais();
    });

    // helper para formatar data para assinatura: YYYY-MM-DD ou DD/MM/YYYY -> "05 de junho de 2025"
    function formatDataParaAssinatura(dateStr) {
        if (!dateStr) return '';
        let d,m,y;
        if (dateStr.includes('-')) { [y,m,d] = dateStr.split('-'); }
        else if (dateStr.includes('/')) { [d,m,y] = dateStr.split('/'); }
        else return dateStr;
        const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
        return `${parseInt(d,10)} de ${meses[parseInt(m,10)-1]} de ${y}`;
    }

    // Montagem do preview (HTML) e preparação para DOCX/PDF
    processarBtn?.addEventListener('click', () => {
        const municipio = (document.getElementById('nome-municipio')?.value || '').trim();
        const prefeito = (document.getElementById('nome-prefeito')?.value || '').trim();
        const numeroPL = document.getElementById('numero-pl')?.value || '___/_____';
        const dataDoc = (document.getElementById('data-pl')?.value || '').trim();
        const justificativa = (document.getElementById('justificativa-pl')?.value || '').trim();
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        const secretariaNome = nomeSecretariaInput?.value || '';
        const secretariaCargo = cargoSecretariaInput?.value || '';

        // coletar itens
        const anulacoes = []; anulacaoContainer.querySelectorAll('.anul-item').forEach(div => {
            const sel = div.querySelector('select'); const inp = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const valor = parseNumberFromString(inp.value) || 0;
            anulacoes.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });
        const creditos = []; creditoContainer.querySelectorAll('.cred-item').forEach(div => {
            const sel = div.querySelector('select'); const inp = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const valor = parseNumberFromString(inp.value) || 0;
            creditos.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });
        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        // montar HTML preview (mais limpo)
        let titulo = tipo === 'decreto' ? 'DECRETO' : (tipo === 'leiFinal' ? 'LEI' : 'PROJETO DE LEI');
        let introLine = tipo === 'decreto' ? 'Dispõe sobre a autorização para abertura de Crédito Adicional Suplementar' : 'Dispõe sobre a abertura de Crédito Adicional Suplementar';

        let html = `<div style="font-family: 'Times New Roman', Times, serif; color:#000; padding:20px; line-height:1.25">`;
        html += `<h2 style="text-align:center; margin-bottom:4px">${titulo}</h2>`;
        html += `<p style="text-align:center; margin-top:0; margin-bottom:8px"><strong>${introLine}</strong></p>`;
        html += `<p style="text-align:center; margin-top:0;"><strong>${municipio}</strong>${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}</p>`;

        if (tipo === 'decreto') {
            html += `<p>O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,</p><p style="text-align:center;"><strong>DECRETA:</strong></p>`;
        } else if (tipo === 'projetoLei') {
            html += `<p>O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:</p>`;
        } else {
            html += `<p>O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:</p>`;
        }

        // Art. 1 - créditos
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s,it)=>s+(it.valor||0),0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            html += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender a(s) seguinte(s) dotação(ões):</p>`;
            html += `<pre style="white-space:pre-wrap; font-family:inherit; font-size:13px;">`;
            creditos.forEach(it => { html += `${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} \t R$ ${formatCurrency(it.valor.toFixed(2))}\n`; });
            html += `\nTOTAL\nR$ ${formatCurrency(somaCred.toFixed(2))}`;
            html += `</pre>`;
        } else {
            html += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.</p>`;
        }

        // Art. 2 - fontes
        const somaAnul = anulacoes.reduce((s,it)=>s+(it.valor||0),0);
        if (anulacoes.length > 0 || superVal || excessoVal) {
            html += `<p><strong>Art. 2º</strong> Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:</p><ul>`;
            if (superVal) html += `<li>Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})</li>`;
            if (excessoVal) html += `<li>Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})</li>`;
            anulacoes.forEach(it => html += `<li>Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}</li>`);
            if (anulacoes.length > 0) html += `<li><strong>TOTAL</strong> R$ ${formatCurrency(somaAnul.toFixed(2))}</li>`;
            html += `</ul>`;
        }

        html += `<p><strong>Art. 3º</strong> As alterações promovidas passam a fazer parte da LDO e do PPA vigentes.</p>`;
        html += `<p><strong>Art. 4º</strong> Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.</p>`;

        html += `<div style="margin-top:30px; text-align:center;">`;
        html += `<p>${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}</p>`;
        html += `<p style="margin-top:30px;"><strong>${prefeito}</strong><br>PREFEITO MUNICIPAL</p>`;
        if (nomeSecretariaInput && nomeSecretariaInput.value) {
            html += `<p style="margin-top:30px;"><strong>${nomeSecretariaInput.value}</strong><br>${(cargoSecretariaInput && cargoSecretariaInput.value) ? cargoSecretariaInput.value : 'Secretaria'}</p>`;
        }
        html += `</div>`;

        if (tipo === 'projetoLei') {
            html += `<div style="page-break-before: always;"></div>`;
            html += `<div style="padding-top:10px;"><h3 style="text-align:center">JUSTIFICATIVA</h3><p style="white-space:pre-wrap; margin-top:6px;">${justificativa}</p></div>`;
        }

        projetoLeiContainer.innerHTML = html;
        projetoLeiContainer.classList.remove('hidden');
        gerarPdfBtn.classList.remove('hidden');
        gerarDocxBtn.classList.remove('hidden');

        calcularTotais();
    });

    // -----------------------------
    // GERAÇÃO DOCX (formatado: Times New Roman 12, espaçamento)
    // -----------------------------
    gerarDocxBtn?.addEventListener('click', async () => {
        if (!window.docx) { alert('Biblioteca docx não carregada.'); return; }
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = window.docx;

        // recuperar mesmo dados do preview (para garantir consistência)
        const municipio = (document.getElementById('nome-municipio')?.value || '').trim();
        const prefeito = (document.getElementById('nome-prefeito')?.value || '').trim();
        const dataDoc = (document.getElementById('data-pl')?.value || '').trim();
        const justificativa = (document.getElementById('justificativa-pl')?.value || '').trim();
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';

        const anulacoes = []; anulacaoContainer.querySelectorAll('.anul-item').forEach(div => {
            const sel = div.querySelector('select'); const inp = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const val = parseNumberFromString(inp.value) || 0;
            anulacoes.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor: val });
        });
        const creditos = []; creditoContainer.querySelectorAll('.cred-item').forEach(div => {
            const sel = div.querySelector('select'); const inp = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const val = parseNumberFromString(inp.value) || 0;
            creditos.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor: val });
        });

        const docChildren = [];

        // Título principal centralizado (Times New Roman 14 bold)
        docChildren.push(new Paragraph({
            children: [ new TextRun({ text: tipo === 'decreto' ? 'DECRETO' : (tipo === 'leiFinal' ? 'LEI' : 'PROJETO DE LEI'), bold: true, size: 28 }) ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 }
        }));

        // linha de introdução central (municipio + data)
        docChildren.push(new Paragraph({
            children: [ new TextRun({ text: `${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, size: 24 }) ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 120 }
        }));

        // corpo intro
        const introText = tipo === 'decreto'
            ? `O Prefeito Municipal de ${municipio}, usando de suas atribuições legais, DECRETA:`
            : (tipo === 'projetoLei'
                ? `O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:`
                : `O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:`);

        docChildren.push(new Paragraph({ children: [ new TextRun({ text: introText, size: 24 }) ], spacing: { after: 120 } }));

        // Artigo 1
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s,it)=>s+(it.valor||0),0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            docChildren.push(new Paragraph({
                children: [ new TextRun({ text: 'Art. 1º ', bold: true, size: 24 }), new TextRun({ text: `Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender a(s) seguinte(s) dotação(ões):`, size: 24 }) ],
                spacing: { after: 80 }
            }));
            // cada crédito em parágrafo monoespaçado
            creditos.forEach(it => {
                docChildren.push(new Paragraph({ children: [ new TextRun({ text: `${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`, size: 24 }) ], spacing: { after: 40 } }));
            });
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: `TOTAL`, bold: true, size: 24 }) ] }));
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: `R$ ${formatCurrency(creditos.reduce((s,it)=>s+(it.valor||0),0).toFixed(2))}`, size: 24 }) , spacing: { after: 80 } }));
        } else {
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 1º ', bold: true, size: 24 }), new TextRun({ text: 'Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.', size: 24 }) ], spacing: { after: 80 } }));
        }

        // Art.2 fontes
        if (anulacoes.length > 0 || parseNumberFromString(valorSuperavitInput?.value) || parseNumberFromString(valorExcessoInput?.value)) {
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 2º ', bold: true, size: 24 }), new TextRun({ text: 'Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:', size: 24 }) ], spacing: { after: 80 } }));
            if (parseNumberFromString(valorSuperavitInput?.value)) docChildren.push(new Paragraph({ children: [ new TextRun({ text: `Superávit Financeiro no valor de R$ ${formatCurrency(parseNumberFromString(valorSuperavitInput.value).toFixed(2))} (${numeroParaExtensoBR(parseNumberFromString(valorSuperavitInput.value).toFixed(2))})`, size: 24 }) ]));
            if (parseNumberFromString(valorExcessoInput?.value)) docChildren.push(new Paragraph({ children: [ new TextRun({ text: `Excesso de Arrecadação no valor de R$ ${formatCurrency(parseNumberFromString(valorExcessoInput.value).toFixed(2))} (${numeroParaExtensoBR(parseNumberFromString(valorExcessoInput.value).toFixed(2))})`, size: 24 }) ]}));
            anulacoes.forEach(it => docChildren.push(new Paragraph({ children: [ new TextRun({ text: `Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`, size: 24 }) ])));
            if (anulacoes.length > 0) docChildren.push(new Paragraph({ children: [ new TextRun({ text: `TOTAL (anulação) R$ ${formatCurrency(anulacoes.reduce((s,it)=>s+(it.valor||0),0).toFixed(2))}`, size: 24 }) ], spacing: { after: 80 })));
        }

        // Art.3 e Art.4
        docChildren.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 3º ', bold: true, size: 24 }), new TextRun({ text: 'As alterações promovidas passam a fazer parte da LDO e do PPA vigentes.', size: 24 }) ], spacing: { after: 80 } }));
        docChildren.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 4º ', bold: true, size: 24 }), new TextRun({ text: `Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.`, size: 24 }) ], spacing: { after: 120 } }));

        // assinatura
        docChildren.push(new Paragraph({ children: [ new TextRun({ text: `${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, size: 24 }) , alignment: AlignmentType.CENTER, spacing: { after: 200 } }));
        docChildren.push(new Paragraph({ children: [ new TextRun({ text: `${prefeito}`, bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }));
        docChildren.push(new Paragraph({ children: [ new TextRun({ text: 'PREFEITO MUNICIPAL', size: 24 }) ], alignment: AlignmentType.CENTER }));

        if (nomeSecretariaInput && nomeSecretariaInput.value) {
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: '', size: 24 }) ] }));
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: `${nomeSecretariaInput.value}`, bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }));
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: `${(cargoSecretariaInput && cargoSecretariaInput.value) ? cargoSecretariaInput.value : 'Secretaria'}`, size: 24 }) ], alignment: AlignmentType.CENTER }));
        }

        // justificativa (nova seção/página se projeto de lei)
        if (document.querySelector('input[name="tipoDocumento"]:checked')?.value === 'projetoLei') {
            docChildren.push(new Paragraph({ children:[ new TextRun({ text: '', size: 24 }) ], pageBreakBefore: true }));
            docChildren.push(new Paragraph({ children: [ new TextRun({ text: 'JUSTIFICATIVA', bold: true, size: 24 }) ], alignment: AlignmentType.CENTER, spacing: { after: 120 } }));
            // dividir justificativa em parágrafos
            (justificativa.split(/\r?\n\r?\n/).filter(Boolean)).forEach(p => {
                docChildren.push(new Paragraph({ children: [ new TextRun({ text: p.trim(), size: 24 }) ], spacing: { after: 120 } }));
            });
        }

        // montar documento com fonte padrão Times New Roman (via run properties)
        // Note: docx lib não permite setar font global facilmente; definimos em cada TextRun via font: 'Times New Roman'
        // ajustar todas TextRuns para terem font 'Times New Roman'
        function withTimes(run) {
            run.font = 'Times New Roman';
            return run;
        }
        docChildren.forEach(p => {
            p.root.forEach?.(() => {}); // no-op, apenas garantia
            // aplicar font em cada TextRun (se houver)
            if (p.options && p.options.children) {
                p.options.children.forEach(ch => { if (ch && ch.properties) ch.properties.font = 'Times New Roman'; });
            } else if (p.children) {
                p.children.forEach(ch => { if (ch) ch.font = 'Times New Roman'; });
            }
        });

        const { Document: DocxDocument, Packer } = window.docx;
        const docxDoc = new DocxDocument({ sections: [{ properties: {}, children: docChildren }] });

        try {
            const blob = await Packer.toBlob(docxDoc);
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `documento_orcamentario.docx`;
            a.click();
        } catch (err) {
            alert('Erro ao gerar DOCX: ' + (err.message || err));
        }
    });

    // -----------------------------
    // GERAÇÃO PDF (jsPDF) com texto selecionável e quebras
    // -----------------------------
    gerarPdfBtn?.addEventListener('click', () => {
        const municipio = (document.getElementById('nome-municipio')?.value || '').trim();
        const prefeito = (document.getElementById('nome-prefeito')?.value || '').trim();
        const dataDoc = (document.getElementById('data-pl')?.value || '').trim();
        const justificativa = (document.getElementById('justificativa-pl')?.value || '').trim();
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';

        const anulacoes = []; anulacaoContainer.querySelectorAll('.anul-item').forEach(div => {
            const sel = div.querySelector('select'); const inp = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const val = parseNumberFromString(inp.value) || 0;
            anulacoes.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor: val });
        });
        const creditos = []; creditoContainer.querySelectorAll('.cred-item').forEach(div => {
            const sel = div.querySelector('select'); const inp = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const val = parseNumberFromString(inp.value) || 0;
            creditos.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor: val });
        });

        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const margin = 18;
        let cursorY = 20;

        pdf.setFont('Times', 'Normal');
        pdf.setFontSize(12);
        const lineHeight = 6.5;

        function addCentered(text, size = 12, bold = false) {
            pdf.setFontSize(size);
            pdf.setFont('Times', bold ? 'Bold' : 'Normal');
            pdf.text(text, pageWidth / 2, cursorY, { align: 'center' });
            cursorY += lineHeight + (size > 12 ? 2 : 0);
        }
        function addParagraph(text, opts = {}) {
            pdf.setFontSize(opts.size || 12);
            pdf.setFont('Times', opts.bold ? 'Bold' : 'Normal');
            const maxw = pageWidth - margin * 2;
            const lines = pdf.splitTextToSize(text, maxw);
            pdf.text(lines, margin, cursorY, { align: 'left' });
            cursorY += lines.length * lineHeight;
            if (cursorY > pageHeight - 30) { pdf.addPage(); cursorY = 20; }
        }

        addCentered(tipo === 'decreto' ? 'DECRETO' : (tipo === 'leiFinal' ? 'LEI' : 'PROJETO DE LEI'), 14, true);
        addCentered('Dispõe sobre a abertura de Crédito Adicional Suplementar', 12, false);
        addCentered(`${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, 12, false);
        cursorY += 4;

        if (tipo === 'decreto') {
            addParagraph(`O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,`);
            addParagraph(`DECRETA:`);
        } else if (tipo === 'projetoLei') {
            addParagraph(`O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:`);
        } else {
            addParagraph(`O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:`);
        }

        // Art.1
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s,it)=>s+(it.valor||0),0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            addParagraph(`Art. 1º Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender a(s) seguinte(s) dotação(ões):`);
            creditos.forEach(it => addParagraph(`${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`));
            addParagraph(`TOTAL\nR$ ${formatCurrency(somaCred.toFixed(2))}`);
        } else {
            addParagraph('Art. 1º Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.');
        }

        // Art.2
        if (anulacoes.length > 0 || superVal || excessoVal) {
            addParagraph('Art. 2º Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:');
            if (superVal) addParagraph(`- Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})`);
            if (excessoVal) addParagraph(`- Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})`);
            anulacoes.forEach(it => addParagraph(`- Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`));
            if (anulacoes.length > 0) addParagraph(`TOTAL\nR$ ${formatCurrency(anulacoes.reduce((s,it)=>s+(it.valor||0),0).toFixed(2))}`);
        }

        addParagraph('Art. 3º As alterações promovidas passam a integrar a LDO e o PPA vigentes.');
        addParagraph(`Art. 4º Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.`);

        // assinatura
        cursorY += 10;
        addParagraph(`${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`);
        cursorY += 10;
        addParagraph(prefeito);
        addParagraph('PREFEITO MUNICIPAL');
        if (nomeSecretariaInput && nomeSecretariaInput.value) {
            cursorY += 10;
            addParagraph(nomeSecretariaInput.value);
            addParagraph((cargoSecretariaInput && cargoSecretariaInput.value) ? cargoSecretariaInput.value : 'Secretaria');
        }

        // justificativa em nova página
        if (document.querySelector('input[name="tipoDocumento"]:checked')?.value === 'projetoLei') {
            pdf.addPage();
            cursorY = 20;
            addCentered('JUSTIFICATIVA', 14, true);
            addParagraph(justificativa || '');
        }

        pdf.save('documento_orcamentario.pdf');
    });

    // loop de segurança para recalcular totais
    setInterval(calcularTotais, 800);
});
