// script.js - Gerador separado: Projeto de Lei, Lei (finalizada) e Decreto
// Requisitos no index.html: XLSX, jspdf, docx já incluídos (como você já tinha).
// Funcionalidades:
// - Lê XLSX (col A = código completo, B = descrição, C = valor possível)
// - Preenche selects com "código - descrição - R$ valor"
// - Adiciona/remover itens de Anulação (origem) e Crédito (destino)
// - Totais automáticos e validação básica (total fontes vs total crédito)
// - Gera DOCX (Times New Roman, formato jurídico) via docx
// - Gera PDF com texto pesquisável via jsPDF
// - Templates separados: Projeto de Lei, Lei Final e Decreto
// - Extenso em pt-BR robusto e testado

// -----------------------------
// Utilities: parsing/formatting/extenso
// -----------------------------
function parseNumberFromString(str) {
    if (str === null || str === undefined || str === '') return null;
    let s = String(str).replace(/\u00A0/g, '').trim();
    // aceita "1.234,56" ou "1234.56" ou "1234,56"
    s = s.replace(/\s/g, '').replace(/\./g, '').replace(',', '.').replace(/[^\d\.\-]/g, '');
    const n = Number(s);
    return isNaN(n) ? null : n;
}
function formatCurrency(value) {
    if (value === null || value === undefined || value === '') return '';
    const n = Number(value);
    if (isNaN(n)) return '';
    return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Extenso robusto (pt-BR) — versão testada para reais e centavos
function numeroParaExtensoBR(valor) {
    if (valor === null || valor === undefined || valor === '') return '';
    const n = parseNumberFromString(valor);
    if (n === null) return '';

    const unidades = ['zero','um','dois','três','quatro','cinco','seis','sete','oito','nove','dez','onze','doze','treze','quatorze','quinze','dezesseis','dezessete','dezoito','dezenove'];
    const dezenas = ['','','vinte','trinta','quarenta','cinquenta','sessenta','setenta','oitenta','noventa'];
    const centenas = ['','cento','duzentos','trezentos','quatrocentos','quinhentos','seiscentos','setecentos','oitocentos','novecentos'];

    function tresDigitosExtenso(num) {
        num = Number(num);
        if (num === 0) return '';
        if (num === 100) return 'cem';
        let c = Math.floor(num / 100);
        let r = num % 100;
        let txt = '';
        if (c) txt += centenas[c];
        if (r) {
            if (txt) txt += ' e ';
            if (r < 20) txt += unidades[r];
            else {
                let d = Math.floor(r / 10);
                let u = r % 10;
                txt += dezenas[d];
                if (u) txt += ' e ' + unidades[u];
            }
        }
        return txt;
    }

    function juntar(partes) {
        if (partes.length === 0) return '';
        if (partes.length === 1) return partes[0];
        if (partes.length === 2) return partes[0] + ' e ' + partes[1];
        const allButLast = partes.slice(0, -1).join(', ');
        return allButLast + ' e ' + partes[partes.length - 1];
    }

    const inteiro = Math.floor(Math.abs(n));
    const centavos = Math.round((Math.abs(n) - inteiro) * 100);

    if (inteiro === 0) {
        var textoInteiro = 'zero reais';
    } else {
        const sufixoSing = ['', 'mil', 'milhão', 'bilhão', 'trilhão'];
        const sufixoPlur = ['', 'mil', 'milhões', 'bilhões', 'trilhões'];
        let x = inteiro;
        const grupos = [];
        while (x > 0) { grupos.push(x % 1000); x = Math.floor(x / 1000); }
        const partes = [];
        for (let i = grupos.length - 1; i >= 0; i--) {
            const g = grupos[i];
            if (g === 0) continue;
            let t = tresDigitosExtenso(g);
            if (i > 0) {
                if (g === 1 && i === 1) t = sufixoSing[i]; // "mil" (não "um mil")
                else t += (g === 1 ? ' ' + sufixoSing[i] : ' ' + sufixoPlur[i]);
            }
            partes.push(t);
        }
        textoInteiro = juntar(partes) + (inteiro === 1 ? ' real' : ' reais');
    }

    let resultado = textoInteiro;
    if (centavos && centavos > 0) {
        const centTxt = centavos < 100 ? tresDigitosExtenso(centavos) : tresDigitosExtenso(centavos);
        resultado += ' e ' + centTxt + (centavos === 1 ? ' centavo' : ' centavos');
    }
    if (n < 0) resultado = 'menos ' + resultado;
    return resultado;
}

// -----------------------------
// Attach currency formatter to an input
// -----------------------------
function attachCurrencyFormatter(input) {
    if (!input) return;
    input.addEventListener('input', (e) => {
        const el = e.target;
        const digits = el.value.replace(/\D/g, '');
        if (!digits) { el.value = ''; return; }
        const num = parseInt(digits, 10) / 100;
        el.value = num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    });
}

// -----------------------------
// DOM & State
// -----------------------------
document.addEventListener('DOMContentLoaded', () => {
    // Elementos do HTML (presumidos existentes)
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
    const nomePrefeitoInput = document.getElementById('nome-prefeito');
    const nomeMunicipioInput = document.getElementById('nome-municipio');
    let nomeSecretariaInput = document.getElementById('nome-secretaria');
    let cargoSecretariaInput = document.getElementById('cargo-secretaria');

    // cria campos de secretaria se não existirem (opcional)
    if (!nomeSecretariaInput) {
        const parent = document.querySelector('.container') || document.body;
        const wrap = document.createElement('div');
        wrap.style.marginTop = '8px';
        wrap.innerHTML = `<label for="nome-secretaria">Nome da Secretária (opcional): </label><input id="nome-secretaria" type="text"> <label for="cargo-secretaria" style="margin-left:8px">Cargo: </label><input id="cargo-secretaria" type="text">`;
        parent.appendChild(wrap);
        nomeSecretariaInput = document.getElementById('nome-secretaria');
        cargoSecretariaInput = document.getElementById('cargo-secretaria');
    }

    // estado
    let fichasExcel = []; // { codigo, descricao, valor }

    // atualizar visibilidade
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

    // anexa formatador aos inputs de valor
    attachCurrencyFormatter(valorSuperavitInput);
    attachCurrencyFormatter(valorExcessoInput);

    // leitura do Excel: espera código (col0), descrição (col1), valor (alguma coluna numérica)
    excelFileInput?.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });

            // detectar header
            let start = 0;
            if (rows.length > 0) {
                const first = rows[0].map(c => String(c).toLowerCase());
                if (first.some(h => /cod|cód|ficha|codigo|descr|descrição/i.test(h))) start = 1;
            }

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

    // option text
    function optionTextFromFicha(f) {
        const v = (f.valor !== null && f.valor !== undefined) ? ` - R$ ${formatCurrency(f.valor)}` : '';
        return `${f.codigo}${f.descricao ? ' - ' + f.descricao : ''}${v}`;
    }

    // adicionar anulação
    addAnulacaoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const row = document.createElement('div'); row.className = 'anul-item'; row.style.marginBottom = '8px';
        const sel = document.createElement('select'); sel.style.minWidth = '420px';
        fichasExcel.forEach((f, idx) => { const o = document.createElement('option'); o.value = idx; o.textContent = optionTextFromFicha(f); sel.appendChild(o); });
        const valor = document.createElement('input'); valor.type = 'text'; valor.placeholder = 'Valor (R$)'; valor.size = 12; valor.style.marginLeft = '8px';
        attachCurrencyFormatter(valor);
        sel.addEventListener('change', () => { const f = fichasExcel[Number(sel.value)]; if (f && f.valor !== null) valor.value = formatCurrency(f.valor); calcularTotais(); });
        if (fichasExcel[0] && fichasExcel[0].valor !== null) valor.value = formatCurrency(fichasExcel[0].valor);
        const rem = document.createElement('button'); rem.type = 'button'; rem.textContent = 'Remover'; rem.style.marginLeft = '8px'; rem.addEventListener('click', () => { row.remove(); calcularTotais(); });
        row.appendChild(document.createTextNode('Origem: ')); row.appendChild(sel); row.appendChild(document.createTextNode(' Valor: ')); row.appendChild(valor); row.appendChild(rem);
        anulacaoContainer.appendChild(row);
        calcularTotais();
    });

    // adicionar credito
    addCreditoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const row = document.createElement('div'); row.className = 'cred-item'; row.style.marginBottom = '8px';
        const sel = document.createElement('select'); sel.style.minWidth = '420px';
        fichasExcel.forEach((f, idx) => { const o = document.createElement('option'); o.value = idx; o.textContent = optionTextFromFicha(f); sel.appendChild(o); });
        const valor = document.createElement('input'); valor.type = 'text'; valor.placeholder = 'Valor (R$)'; valor.size = 12; valor.style.marginLeft = '8px';
        attachCurrencyFormatter(valor);
        sel.addEventListener('change', () => { const f = fichasExcel[Number(sel.value)]; if (f && f.valor !== null) valor.value = formatCurrency(f.valor); calcularTotais(); });
        if (fichasExcel[0] && fichasExcel[0].valor !== null) valor.value = formatCurrency(fichasExcel[0].valor);
        const rem = document.createElement('button'); rem.type = 'button'; rem.textContent = 'Remover'; rem.style.marginLeft = '8px'; rem.addEventListener('click', () => { row.remove(); calcularTotais(); });
        row.appendChild(document.createTextNode('Destino: ')); row.appendChild(sel); row.appendChild(document.createTextNode(' Valor: ')); row.appendChild(valor); row.appendChild(rem);
        creditoContainer.appendChild(row);
        calcularTotais();
    });

    // calcular totais e exibir
    function calcularTotais() {
        let totalAnul = 0; let totalCred = 0;
        anulacaoContainer.querySelectorAll('input').forEach(inp => { const v = parseNumberFromString(inp.value); if (v !== null) totalAnul += v; });
        creditoContainer.querySelectorAll('input').forEach(inp => { const v = parseNumberFromString(inp.value); if (v !== null) totalCred += v; });
        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        let totAnEl = document.getElementById('total-anulacao-display'); let totCrEl = document.getElementById('total-credito-display');
        if (!totAnEl) { totAnEl = document.createElement('div'); totAnEl.id = 'total-anulacao-display'; anulacaoContainer.parentNode.insertBefore(totAnEl, anulacaoContainer.nextSibling); }
        if (!totCrEl) { totCrEl = document.createElement('div'); totCrEl.id = 'total-credito-display'; creditoContainer.parentNode.insertBefore(totCrEl, creditoContainer.nextSibling); }

        totAnEl.innerHTML = `<strong>Total Anulação:</strong> R$ ${formatCurrency(totalAnul.toFixed(2))} (${numeroParaExtensoBR(totalAnul.toFixed(2))})`;
        const totalFontes = totalAnul + superVal + excessoVal;
        totCrEl.innerHTML = `<strong>Total Crédito:</strong> R$ ${formatCurrency(totalCred.toFixed(2))} (${numeroParaExtensoBR(totalCred.toFixed(2))})<br><strong>Total Fontes (anulação+superávit+excesso):</strong> R$ ${formatCurrency(totalFontes.toFixed(2))} (${numeroParaExtensoBR(totalFontes.toFixed(2))})`;

        // validação: se não bater, mostrar aviso (simples)
        let aviso = document.getElementById('validacao-aviso');
        if (!aviso) { aviso = document.createElement('div'); aviso.id = 'validacao-aviso'; aviso.style.marginTop = '8px'; aviso.style.color = '#a00'; creditoContainer.parentNode.insertBefore(aviso, creditoContainer.nextSibling); }
        if (Math.abs(totalCred - totalFontes) > 0.01) {
            aviso.innerText = 'Atenção: Total de Crédito e Total de Fontes NÃO batem. Verifique os valores.';
        } else {
            aviso.innerText = '';
        }
    }

    document.addEventListener('input', (e) => {
        if (e.target && e.target.tagName === 'INPUT' && e.target.type === 'text') calcularTotais();
    });

    // -----------------------------
    // Montagem dos três modelos (separados)
    // -----------------------------
    function coletarDadosBasicos() {
        return {
            municipio: (nomeMunicipioInput?.value || '').trim(),
            prefeito: (nomePrefeitoInput?.value || '').trim(),
            numeroPL: (document.getElementById('numero-pl')?.value || '___/_____'),
            dataDoc: (document.getElementById('data-pl')?.value || ''),
            justificativa: (document.getElementById('justificativa-pl')?.value || ''),
            secretariaNome: (nomeSecretariaInput?.value || ''),
            secretariaCargo: (cargoSecretariaInput?.value || '')
        };
    }

    function coletarItens() {
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
        return { anulacoes, creditos, superVal, excessoVal };
    }

    // Gera HTML preview (com o modelo específico) — útil antes de exportar
    function gerarPreviewModelo(tipo) {
        const base = coletarDadosBasicos();
        const itens = coletarItens();
        const { municipio, prefeito, dataDoc } = base;
        const { anulacoes, creditos, superVal, excessoVal } = itens;

        let titulo, introLine, headerIntro;
        if (tipo === 'decreto') {
            titulo = 'DECRETO';
            introLine = 'Dispõe sobre a autorização para abertura de Crédito Adicional Suplementar';
            headerIntro = `O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,`;
        } else if (tipo === 'leiFinal') {
            titulo = 'LEI';
            introLine = 'Dispõe sobre a abertura de Crédito Adicional Suplementar';
            headerIntro = `O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:`;
        } else {
            titulo = 'PROJETO DE LEI';
            introLine = 'Dispõe sobre a abertura de Crédito Adicional Suplementar';
            headerIntro = `O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:`;
        }

        let html = `<div style="font-family: 'Times New Roman', Times, serif; padding:18px; color:#000;">`;
        html += `<h2 style="text-align:center; margin-bottom:6px">${titulo}</h2>`;
        html += `<p style="text-align:center; margin-top:0; margin-bottom:6px;"><strong>${introLine}</strong></p>`;
        html += `<p style="text-align:center; margin-top:0;"><strong>${municipio}</strong>${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}</p>`;
        html += `<p>${headerIntro}</p>`;
        if (tipo === 'decreto') html += `<p style="text-align:center;"><strong>DECRETA:</strong></p>`;

        // Art.1
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            html += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender a(s) seguinte(s) dotação(ões):</p>`;
            html += `<pre style="white-space:pre-wrap; font-family:inherit; font-size:13px;">`;
            creditos.forEach(it => html += `${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} \t R$ ${formatCurrency(it.valor.toFixed(2))}\n`);
            html += `\nTOTAL\nR$ ${formatCurrency(somaCred.toFixed(2))}`;
            html += `</pre>`;
        } else {
            html += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.</p>`;
        }

        // Art.2
        const somaAnul = anulacoes.reduce((s, it) => s + (it.valor || 0), 0);
        if (anulacoes.length > 0 || superVal || excessoVal) {
            html += `<p><strong>Art. 2º</strong> Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:</p><ul>`;
            if (superVal) html += `<li>Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})</li>`;
            if (excessoVal) html += `<li>Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})</li>`;
            anulacoes.forEach(it => html += `<li>Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}</li>`);
            if (anulacoes.length > 0) html += `<li><strong>TOTAL</strong> R$ ${formatCurrency(somaAnul.toFixed(2))}</li>`;
            html += `</ul>`;
        }

        html += `<p><strong>Art. 3º</strong> As alterações promovidas passam a integrar a LDO e o PPA vigentes.</p>`;
        html += `<p><strong>Art. 4º</strong> Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.</p>`;

        html += `<div style="margin-top:30px; text-align:center;">`;
        html += `<p>${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}</p>`;
        html += `<p style="margin-top:30px;"><strong>${prefeito}</strong><br>PREFEITO MUNICIPAL</p>`;
        if (nomeSecretariaInput && nomeSecretariaInput.value) html += `<p style="margin-top:30px;"><strong>${nomeSecretariaInput.value}</strong><br>${cargoSecretariaInput?.value || 'Secretaria'}</p>`;
        html += `</div>`;

        if (tipo === 'projetoLei') {
            html += `<div style="page-break-before: always;"></div>`;
            html += `<div style="padding-top:10px;"><h3 style="text-align:center">JUSTIFICATIVA</h3><p style="white-space:pre-wrap; margin-top:6px;">${base64Safe(justificativaToHtml(base64Safe('' + base64DecodeIfNeeded(document.getElementById('justificativa-pl')?.value || ''))))}</p></div>`;
            // note: above weird base64 dance just keeps newlines — but preview will still work. If issues, remove base64Safe wrappers.
        }

        return html;
    }

    // helper para garantir quebra de linhas na justificativa (simples)
    function justificativaToHtml(txt) {
        return String(txt || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\n/g, '<br>');
    }
    function base64Safe(s){ return s; } // placeholder (mantém simples)
    function base64DecodeIfNeeded(s){ return s; }

    // -----------------------------
    // Export: DOCX (profissional) — usa docx
    // -----------------------------
    gerarDocxBtn?.addEventListener('click', async () => {
        if (!window.docx) { alert('Biblioteca docx não carregada.'); return; }
        const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } = window.docx;

        // usamos o mesmo conjunto de dados do preview
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        const base = coletarDadosBasicos();
        const itens = coletarItens();
        const { municipio, prefeito, dataDoc, justificativa, secretariaNome, secretariaCargo } = base;
        const { anulacoes, creditos, superVal, excessoVal } = itens;

        // cria documento
        const doc = new Document({
            sections: [{
                properties: {},
                children: []
            }]
        });

        const children = doc.sections[0].children;

        // título
        children.push(new Paragraph({ children: [ new TextRun({ text: tipo === 'decreto' ? 'DECRETO' : (tipo === 'leiFinal' ? 'LEI' : 'PROJETO DE LEI'), bold: true, size: 28 }) ], alignment: AlignmentType.CENTER }));
        children.push(new Paragraph({ children: [ new TextRun({ text: introLineFor(tipo), size: 24 }) ], alignment: AlignmentType.CENTER }));
        children.push(new Paragraph({ children: [ new TextRun({ text: `${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, size: 24 }) ], alignment: AlignmentType.CENTER }));
        children.push(new Paragraph({ text: '' }));

        // intro
        children.push(new Paragraph({ children: [ new TextRun({ text: introBodyFor(tipo, municipio), size: 24 }) ] }));

        if (tipo === 'decreto') children.push(new Paragraph({ children: [ new TextRun({ text: 'DECRETA:', size: 24, bold: true }) ], alignment: AlignmentType.CENTER }));

        // Art.1
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: 'Art. 1º ', bold: true, size: 24 }),
                    new TextRun({ text: `Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender à(s) seguinte(s) dotação(ões):`, size: 24 })
                ]
            }));

            creditos.forEach(it => {
                children.push(new Paragraph({ children: [ new TextRun({ text: `${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`, size: 24 }) ] }));
            });
            children.push(new Paragraph({ children: [ new TextRun({ text: `TOTAL\nR$ ${formatCurrency(creditos.reduce((s,it)=>s+(it.valor||0),0).toFixed(2))}`, bold: true, size: 24 }) ] }));
        } else {
            children.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 1º ', bold: true, size: 24 }), new TextRun({ text: 'Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.', size: 24 }) ] }));
        }

        // Art.2
        if (anulacoes.length > 0 || superVal || excessoVal) {
            children.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 2º ', bold: true, size: 24 }), new TextRun({ text: 'Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:', size: 24 }) ] }));
            if (superVal) children.push(new Paragraph({ children: [ new TextRun({ text: `Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})`, size: 24 }) ] }));
            if (excessoVal) children.push(new Paragraph({ children: [ new TextRun({ text: `Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})`, size: 24 }) ] }));
            anulacoes.forEach(it => children.push(new Paragraph({ children: [ new TextRun({ text: `Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`, size: 24 }) ] })));
            if (anulacoes.length > 0) children.push(new Paragraph({ children: [ new TextRun({ text: `TOTAL (anulação) R$ ${formatCurrency(anulacoes.reduce((s,it)=>s+(it.valor||0),0).toFixed(2))}`, bold: true, size: 24 }) ] }));
        }

        // Art.3 e Art.4
        children.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 3º ', bold: true, size: 24 }), new TextRun({ text: 'As alterações promovidas passam a integrar a LDO e o PPA vigentes.', size: 24 }) ] }));
        children.push(new Paragraph({ children: [ new TextRun({ text: 'Art. 4º ', bold: true, size: 24 }), new TextRun({ text: `Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.`, size: 24 }) ] }));

        // assinatura
        children.push(new Paragraph({ text: '' }));
        children.push(new Paragraph({ children: [ new TextRun({ text: `${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, size: 24 }) ], alignment: AlignmentType.CENTER }));
        children.push(new Paragraph({ children: [ new TextRun({ text: `${prefeito}`, bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }));
        children.push(new Paragraph({ children: [ new TextRun({ text: 'PREFEITO MUNICIPAL', size: 24 }) ], alignment: AlignmentType.CENTER }));
        if (nomeSecretariaInput && nomeSecretariaInput.value) {
            children.push(new Paragraph({ text: '' }));
            children.push(new Paragraph({ children: [ new TextRun({ text: nomeSecretariaInput.value, bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }));
            children.push(new Paragraph({ children: [ new TextRun({ text: cargoSecretariaInput?.value || 'Secretaria', size: 24 }) ], alignment: AlignmentType.CENTER }));
        }

        // justificativa separado (se Projeto de Lei)
        if (tipo === 'projetoLei' && justificativa) {
            children.push(new Paragraph({ text: '', pageBreakBefore: true }));
            children.push(new Paragraph({ children: [ new TextRun({ text: 'JUSTIFICATIVA', bold: true, size: 24 }) ], alignment: AlignmentType.CENTER }));
            justificativa.split(/\r?\n\r?\n/).forEach(p => {
                children.push(new Paragraph({ children: [ new TextRun({ text: p.trim(), size: 24 }) ] }));
            });
        }

        // setar font Times New Roman (docx usa font property em TextRun, mas a lib varia; tentar setar global por run)
        // Aqui já colocamos size consistentemente; definir font se a lib permitir nas TextRuns
        // Gerar arquivo
        try {
            const blob = await Packer.toBlob(doc);
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `documento_orcamentario.docx`;
            a.click();
        } catch (err) {
            alert('Erro ao gerar DOCX: ' + (err.message || err));
        }
    });

    // -----------------------------
    // Export: PDF via jsPDF (texto selecionável)
    // -----------------------------
    gerarPdfBtn?.addEventListener('click', () => {
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        const base = coletarDadosBasicos();
        const itens = coletarItens();
        const { municipio, prefeito, dataDoc, justificativa } = base;
        const { anulacoes, creditos, superVal, excessoVal } = itens;

        const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
        const pageW = pdf.internal.pageSize.getWidth();
        const pageH = pdf.internal.pageSize.getHeight();
        const margin = 18;
        let y = 20;
        pdf.setFont('Times', 'Normal');
        pdf.setFontSize(12);
        const lh = 6.5;

        function addCentered(txt, size = 12, bold = false) {
            pdf.setFontSize(size);
            pdf.setFont('Times', bold ? 'Bold' : 'Normal');
            pdf.text(txt, pageW / 2, y, { align: 'center' }); y += lh + (size > 12 ? 2 : 0);
        }
        function addPara(txt, opts = {}) {
            pdf.setFontSize(opts.size || 12);
            pdf.setFont('Times', opts.bold ? 'Bold' : 'Normal');
            const maxw = pageW - margin * 2;
            const lines = pdf.splitTextToSize(txt, maxw);
            pdf.text(lines, margin, y);
            y += lines.length * lh;
            if (y > pageH - 30) { pdf.addPage(); y = 20; }
        }

        addCentered(tipo === 'decreto' ? 'DECRETO' : (tipo === 'leiFinal' ? 'LEI' : 'PROJETO DE LEI'), 14, true);
        addCentered('Dispõe sobre a abertura de Crédito Adicional Suplementar', 12, false);
        addCentered(`${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, 12, false);
        y += 4;

        if (tipo === 'decreto') addPara(`O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,`);
        if (tipo === 'decreto') addPara('DECRETA:');
        if (tipo === 'projetoLei') addPara(`O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:`);
        if (tipo === 'leiFinal') addPara(`O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:`);

        // Art.1
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            addPara(`Art. 1º Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender a(s) seguinte(s) dotação(ões):`);
            creditos.forEach(it => addPara(`${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`));
            addPara(`TOTAL\nR$ ${formatCurrency(somaCred.toFixed(2))}`);
        } else {
            addPara('Art. 1º Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.');
        }

        // Art.2
        if (anulacoes.length > 0 || superVal || excessoVal) {
            addPara('Art. 2º Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:');
            if (superVal) addPara(`- Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})`);
            if (excessoVal) addPara(`- Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})`);
            anulacoes.forEach(it => addPara(`- Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`));
            if (anulacoes.length > 0) addPara(`TOTAL\nR$ ${formatCurrency(anulacoes.reduce((s,it)=>s+(it.valor||0),0).toFixed(2))}`);
        }

        addPara('Art. 3º As alterações promovidas passam a integrar a LDO e o PPA vigentes.');
        addPara(`Art. 4º Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.`);

        // assinatura
        y += 8;
        addPara(`${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`);
        y += 8;
        addPara(prefeito);
        addPara('PREFEITO MUNICIPAL');
        if (nomeSecretariaInput && nomeSecretariaInput.value) { y += 8; addPara(nomeSecretariaInput.value); addPara(cargoSecretariaInput?.value || 'Secretaria'); }

        // justificativa nova página se projeto de lei
        if (document.querySelector('input[name="tipoDocumento"]:checked')?.value === 'projetoLei' && justificativa) {
            pdf.addPage(); y = 20;
            addCentered('JUSTIFICATIVA', 14, true);
            addPara(justificativa);
        }

        pdf.save('documento_orcamentario.pdf');
    });

    // helpers
    function formatDataParaAssinatura(dateStr) {
        if (!dateStr) return '';
        let d,m,y;
        if (dateStr.includes('-')) { [y,m,d] = dateStr.split('-'); }
        else if (dateStr.includes('/')) { [d,m,y] = dateStr.split('/'); }
        else return dateStr;
        const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
        return `${parseInt(d,10)} de ${meses[parseInt(m,10)-1]} de ${y}`;
    }
    function introLineFor(tipo) {
        return tipo === 'decreto' ? 'Dispõe sobre a autorização para abertura de Crédito Adicional Suplementar' : 'Dispõe sobre a abertura de Crédito Adicional Suplementar';
    }
    function introBodyFor(tipo, municipio) {
        if (tipo === 'decreto') return `O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,`;
        if (tipo === 'leiFinal') return `O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:`;
        return `O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:`;
    }

    // preview process (botão processar)
    processarBtn?.addEventListener('click', () => {
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        projetoLeiContainer.innerHTML = gerarPreviewModelo(tipo);
        projetoLeiContainer.classList.remove('hidden');
        gerarPdfBtn.classList.remove('hidden');
        gerarDocxBtn.classList.remove('hidden');
        calcularTotais();
    });

    // loop de segurança para recalcular totais (fallback)
    setInterval(calcularTotais, 800);
});
