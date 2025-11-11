// script.js - Gerador oficial (PL / LEI / DECRETO) com DOCX e PDF fiéis aos modelos enviados.
// Requisitos: index.html já deve carregar XLSX, jspdf, docx (como no seu index.html anterior).

// -----------------------------
// Helpers: formatação e extenso
// -----------------------------
function formatCurrency(value) {
    if (value === null || value === undefined || value === '') return '';
    const n = Number(String(value).toString().replace(/\s/g, '').replace(/\./g, '').replace(',', '.'));
    if (isNaN(n)) return '';
    return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseNumberFromString(str) {
    if (str === null || str === undefined || str === '') return null;
    const cleaned = String(str).replace(/\s/g, '').replace(/\./g, '').replace(',', '.').replace(/[^\d\.-]/g, '');
    const n = Number(cleaned);
    return isNaN(n) ? null : n;
}

// extenso em pt-BR (valores até trilhões, simplicado porém robusto)
function numeroParaExtensoBR(valor) {
    if (valor === null || valor === undefined || valor === '') return '';
    const n = parseNumberFromString(valor);
    if (n === null) return '';

    const unidades = ['zero','um','dois','três','quatro','cinco','seis','sete','oito','nove','dez','onze','doze','treze','quatorze','quinze','dezesseis','dezessete','dezoito','dezenove'];
    const dezenas = ['','','vinte','trinta','quarenta','cinquenta','sessenta','setenta','oitenta','noventa'];
    const centenas = ['','cento','duzentos','trezentos','quatrocentos','quinhentos','seiscentos','setecentos','oitocentos','novecentos'];

    function tresDigitosParaExtenso(num) {
        num = Number(num);
        if (num === 0) return '';
        if (num === 100) return 'cem';
        let c = Math.floor(num / 100);
        let r = num % 100;
        let texto = '';
        if (c) texto += centenas[c];
        if (r) {
            if (texto) texto += ' e ';
            if (r < 20) texto += unidades[r];
            else {
                let d = Math.floor(r / 10);
                let u = r % 10;
                texto += dezenas[d];
                if (u) texto += ' e ' + unidades[u];
            }
        }
        return texto;
    }

    let inteiro = Math.floor(Math.abs(n));
    let centavos = Math.round((Math.abs(n) - inteiro) * 100);

    if (inteiro === 0) {
        var parteInteiro = 'zero reais';
    } else {
        const nomes = ['', 'mil', 'milhão', 'bilhão', 'trilhão'];
        const nomesPlural = ['', 'mil', 'milhões', 'bilhões', 'trilhões'];
        const grupos = [];
        while (inteiro > 0) {
            grupos.push(inteiro % 1000);
            inteiro = Math.floor(inteiro / 1000);
        }
        const partes = [];
        for (let i = grupos.length - 1; i >= 0; i--) {
            const g = grupos[i];
            if (g === 0) continue;
            let textoGrupo = tresDigitosParaExtenso(g);
            let sufixo = '';
            if (i > 0) {
                sufixo = (g === 1 ? ' ' + nomes[i] : ' ' + nomesPlural[i]);
            }
            partes.push((textoGrupo + sufixo).trim());
        }
        parteInteiro = partes.join(' e ') + ' reais';
    }

    let resultado = parteInteiro;
    if (centavos && centavos > 0) {
        const cExt = tresDigitosParaExtenso(centavos);
        resultado += ' e ' + cExt + (centavos === 1 ? ' centavo' : ' centavos');
    }
    if (n < 0) resultado = 'menos ' + resultado;
    return resultado;
}

// -----------------------------
// Small DOM utilities
// -----------------------------
function ensureInputExists(id, labelText, parentSelector = 'body') {
    let el = document.getElementById(id);
    if (!el) {
        const parent = document.querySelector(parentSelector) || document.body;
        const wrapper = document.createElement('div');
        wrapper.style.marginTop = '6px';
        const label = document.createElement('label');
        label.textContent = labelText + ': ';
        el = document.createElement('input');
        el.id = id;
        el.type = 'text';
        wrapper.appendChild(label);
        wrapper.appendChild(el);
        parent.appendChild(wrapper);
    }
    return el;
}

// -----------------------------
// Main
// -----------------------------
document.addEventListener('DOMContentLoaded', () => {
    // elementos existentes no seu index.html
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

    // campos adicionais que o usuário confirmou (assinaturas)
    const nomePrefeitoInput = document.getElementById('nome-prefeito') || ensureInputExists('nome-prefeito', 'Nome do Prefeito', '.container');
    const nomeMunicipioInput = document.getElementById('nome-municipio') || ensureInputExists('nome-municipio', 'Nome do Município', '.container');
    const dataDocumentoInput = document.getElementById('data-pl') || ensureInputExists('data-pl', 'Data do Documento', '.container');
    // campo para secretária / segunda assinatura
    let nomeSecretariaInput = document.getElementById('nome-secretaria');
    if (!nomeSecretariaInput) {
        // criar campo opcional para secretaria/cargo
        const parent = document.querySelector('.container') || document.body;
        const wrap = document.createElement('div');
        wrap.style.marginTop = '8px';
        wrap.innerHTML = `<label for="nome-secretaria">Nome da Secretária (opcional): </label><input id="nome-secretaria" type="text"> <label for="cargo-secretaria" style="margin-left:8px">Cargo: </label><input id="cargo-secretaria" type="text">`;
        parent.appendChild(wrap);
        nomeSecretariaInput = document.getElementById('nome-secretaria');
    }
    const cargoSecretariaInput = document.getElementById('cargo-secretaria') || document.getElementById('cargo-secretaria');

    // garantia: botões existem
    if (!processarBtn || !gerarPdfBtn || !gerarDocxBtn) {
        console.warn('Botões processar/gerar não encontrados no HTML.');
    }

    // dados das fichas lidas do Excel
    let fichasExcel = []; // { codigo, descricao, valor }

    // atualizar visibilidade e habilitação
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

    // leitura do Excel - espera código (col0), descricao (col1), valor (qualquer coluna numérica depois)
    excelFileInput?.addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function (evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });

            // detectar cabeçalho (se a primeira linha parece conter "cod" ou "ficha")
            let start = 0;
            if (rows.length > 0) {
                const first = rows[0].map(c => String(c).toLowerCase());
                if (first.some(cell => /cod|cód|ficha|codigo|descri|descrição|descricao/i.test(cell))) start = 1;
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
                // se não encontrou, testar coluna 2 mesmo como string numérica
                if (valor === null && r[2] !== undefined) valor = parseNumberFromString(r[2]);

                if (codigo || descricao) {
                    fichasExcel.push({ codigo: codigo || descricao, descricao: descricao || '', valor: valor });
                }
            }

            fichasCount.textContent = `Fichas carregadas: ${fichasExcel.length}`;
            excelDisplay?.classList.remove('hidden');
            atualizarVisibilidade();
        };
        reader.readAsArrayBuffer(file);
    });

    // helper: option text from ficha (exibe o código inteiro como solicitado)
    function optionTextFromFicha(f) {
        const v = (f.valor !== null && f.valor !== undefined) ? ` - R$ ${formatCurrency(f.valor)}` : '';
        return `${f.codigo}${f.descricao ? ' - ' + f.descricao : ''}${v}`;
    }

    // criar item de anulação (origem)
    addAnulacaoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const row = document.createElement('div');
        row.className = 'anul-item';
        row.style.marginBottom = '8px';

        const sel = document.createElement('select');
        sel.style.minWidth = '420px';
        fichasExcel.forEach((f, idx) => {
            const o = document.createElement('option');
            o.value = idx;
            o.textContent = optionTextFromFicha(f);
            sel.appendChild(o);
        });

        const valor = document.createElement('input');
        valor.type = 'text';
        valor.placeholder = 'Valor (R$)';
        valor.size = 12;
        valor.style.marginLeft = '8px';
        valor.addEventListener('input', (e) => {
            const el = e.target;
            const digits = el.value.replace(/\D/g, '');
            if (!digits) { el.value = ''; return; }
            const num = parseInt(digits, 10) / 100;
            el.value = num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            calcularTotais();
        });

        // preencher valor se ficha tiver valor
        if (fichasExcel[0] && fichasExcel[0].valor !== null) {
            valor.value = formatCurrency(fichasExcel[0].valor);
        }
        sel.addEventListener('change', () => {
            const f = fichasExcel[Number(sel.value)];
            if (f && f.valor !== null) valor.value = formatCurrency(f.valor);
            calcularTotais();
        });

        const rem = document.createElement('button');
        rem.type = 'button';
        rem.textContent = 'Remover';
        rem.style.marginLeft = '8px';
        rem.addEventListener('click', () => { row.remove(); calcularTotais(); });

        row.appendChild(document.createTextNode('Origem: '));
        row.appendChild(sel);
        row.appendChild(document.createTextNode(' Valor: '));
        row.appendChild(valor);
        row.appendChild(rem);

        anulacaoContainer.appendChild(row);
        calcularTotais();
    });

    // criar item de crédito (destino)
    addCreditoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const row = document.createElement('div');
        row.className = 'cred-item';
        row.style.marginBottom = '8px';

        const sel = document.createElement('select');
        sel.style.minWidth = '420px';
        fichasExcel.forEach((f, idx) => {
            const o = document.createElement('option');
            o.value = idx;
            o.textContent = optionTextFromFicha(f);
            sel.appendChild(o);
        });

        const valor = document.createElement('input');
        valor.type = 'text';
        valor.placeholder = 'Valor (R$)';
        valor.size = 12;
        valor.style.marginLeft = '8px';
        valor.addEventListener('input', (e) => {
            const el = e.target;
            const digits = el.value.replace(/\D/g, '');
            if (!digits) { el.value = ''; return; }
            const num = parseInt(digits, 10) / 100;
            el.value = num.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            calcularTotais();
        });

        sel.addEventListener('change', () => { const f = fichasExcel[Number(sel.value)]; if (f && f.valor !== null) valor.value = formatCurrency(f.valor); calcularTotais(); });

        const rem = document.createElement('button');
        rem.type = 'button';
        rem.textContent = 'Remover';
        rem.style.marginLeft = '8px';
        rem.addEventListener('click', () => { row.remove(); calcularTotais(); });

        row.appendChild(document.createTextNode('Destino: '));
        row.appendChild(sel);
        row.appendChild(document.createTextNode(' Valor: '));
        row.appendChild(valor);
        row.appendChild(rem);

        creditoContainer.appendChild(row);
        calcularTotais();
    });

    // calcular totais e mostrar abaixo dos containers
    function calcularTotais() {
        let totalAnul = 0;
        let totalCred = 0;

        anulacaoContainer.querySelectorAll('input').forEach(inp => {
            const n = parseNumberFromString(inp.value);
            if (n !== null) totalAnul += n;
        });
        creditoContainer.querySelectorAll('input').forEach(inp => {
            const n = parseNumberFromString(inp.value);
            if (n !== null) totalCred += n;
        });

        // incluir superávit/excesso
        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        // mostrar
        let totAnEl = document.getElementById('total-anulacao-display');
        let totCrEl = document.getElementById('total-credito-display');
        if (!totAnEl) {
            totAnEl = document.createElement('div'); totAnEl.id = 'total-anulacao-display'; anulacaoContainer.parentNode.insertBefore(totAnEl, anulacaoContainer.nextSibling);
        }
        if (!totCrEl) {
            totCrEl = document.createElement('div'); totCrEl.id = 'total-credito-display'; creditoContainer.parentNode.insertBefore(totCrEl, creditoContainer.nextSibling);
        }
        totAnEl.innerHTML = `<strong>Total Anulação:</strong> R$ ${formatCurrency(totalAnul.toFixed(2))} (${numeroParaExtensoBR(totalAnul.toFixed(2))})`;
        const totalFontes = totalAnul + superVal + excessoVal;
        totCrEl.innerHTML = `<strong>Total Crédito:</strong> R$ ${formatCurrency(totalCred.toFixed(2))} (${numeroParaExtensoBR(totalCred.toFixed(2))}) <br>
                             <strong>Total Fontes (anulação+superávit+excesso):</strong> R$ ${formatCurrency(totalFontes.toFixed(2))} (${numeroParaExtensoBR(totalFontes.toFixed(2))})`;
    }

    // recalcular quando qualquer input de moeda mudar (delegation)
    document.addEventListener('input', (e) => {
        if (e.target && e.target.tagName === 'INPUT' && e.target.type === 'text') {
            calcularTotais();
        }
    });

    // monta o texto do documento conforme modelo e gera preview (HTML)
    processarBtn?.addEventListener('click', () => {
        const municipio = (document.getElementById('nome-municipio')?.value || nomeMunicipioInput.value || '').trim();
        const prefeito = (document.getElementById('nome-prefeito')?.value || nomePrefeitoInput.value || '').trim();
        const numeroPL = document.getElementById('numero-pl')?.value || '___/_____';
        const dataDoc = (document.getElementById('data-pl')?.value || dataDocumentoInput.value || '').trim();
        const justificativa = (document.getElementById('justificativa-pl')?.value || '').trim();
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        const secretariaNome = nomeSecretariaInput?.value || '';
        const secretariaCargo = cargoSecretariaInput?.value || '';

        // coletar itens
        const anulacoes = [];
        anulacaoContainer.querySelectorAll('.anul-item').forEach(div => {
            const sel = div.querySelector('select');
            const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)];
            const valor = parseNumberFromString(input.value) || 0;
            anulacoes.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });
        const creditos = [];
        creditoContainer.querySelectorAll('.cred-item').forEach(div => {
            const sel = div.querySelector('select');
            const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)];
            const valor = parseNumberFromString(input.value) || 0;
            creditos.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });

        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        // montar HTML (preview)
        let titulo = tipo === 'projetoLei' ? 'PROJETO DE LEI' : (tipo === 'leiFinal' ? 'LEI' : 'DECRETO');
        let headerLine = tipo === 'decreto' ? 'Dispõe sobre a autorização para abertura de Crédito Adicional Suplementar' : 'Dispõe sobre a abertura de Crédito Adicional Suplementar';

        let html = `<div style="font-family: Arial, sans-serif; padding: 18px; color:#000;">`;
        // topo com Nº e data (se for decreto, mantém padrão do seu exemplo)
        html += `<h3 style="text-align:center; margin-bottom:8px;">${titulo}</h3>`;
        html += `<p style="text-align:center; margin-top:0;"><strong>${headerLine}</strong></p>`;
        html += `<p style="text-align:center;"><strong>${municipio}</strong>${dataDoc ? ', ' + dataDoc : ''}</p>`;

        if (tipo === 'decreto') {
            html += `<p>O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,</p>`;
            html += `<p style="text-align:center;"><strong>DECRETA:</strong></p>`;
        } else if (tipo === 'projetoLei') {
            html += `<p>O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:</p>`;
        } else {
            html += `<p>O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:</p>`;
        }

        // Art.1: créditos (lista de destinos)
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            html += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender à(s) seguinte(s) dotação(ões):</p>`;
            html += `<pre style="white-space:pre-wrap; font-family: inherit; font-size: 14px;">`;
            creditos.forEach(it => {
                html += `${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} \t R$ ${formatCurrency(it.valor.toFixed(2))}\n`;
            });
            html += `\nTOTAL\nR$ ${formatCurrency(somaCred.toFixed(2))}`;
            html += `</pre>`;
        } else {
            html += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.</p>`;
        }

        // Art.2: cobertura (anulações / superávit / excesso)
        const somaAnul = anulacoes.reduce((s, it) => s + (it.valor || 0), 0);
        if (anulacoes.length > 0 || superVal || excessoVal) {
            html += `<p><strong>Art. 2º</strong> Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:</p>`;
            html += `<ul>`;
            if (superVal) html += `<li>Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})</li>`;
            if (excessoVal) html += `<li>Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})</li>`;
            anulacoes.forEach(it => html += `<li>Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}</li>`);
            if (anulacoes.length > 0) html += `<li><strong>TOTAL</strong> R$ ${formatCurrency(somaAnul.toFixed(2))}</li>`;
            html += `</ul>`;
        }

        // Art.3 e Art.4 (padronizados)
        html += `<p><strong>Art. 3º</strong> As alterações promovidas nos artigos anteriores passam a fazer parte da LDO e do PPA vigentes, para fins do disposto em legislação aplicável.</p>`;
        html += `<p><strong>Art. 4º</strong> Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.</p>`;

        // assinatura / rodapé
        html += `<div style="margin-top:30px; text-align:center;">`;
        html += `<p>${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}</p>`;
        html += `<p style="margin-top:40px;"><strong>${prefeito}</strong><br>PREFEITO MUNICIPAL</p>`;
        if (secretariaNome) {
            html += `<p style="margin-top:30px;"><strong>${secretariaNome}</strong><br>${secretariaCargo || 'Secretaria'}</p>`;
        }
        html += `</div>`;

        // justificativa em página separada, se for projeto de lei
        if (tipo === 'projetoLei') {
            html += `<div style="page-break-before: always;"></div>`;
            html += `<div style="padding:18px;"><h3 style="text-align:center">JUSTIFICATIVA</h3>`;
            html += `<p style="white-space:pre-wrap;">${justificativa}</p></div>`;
        }

        projetoLeiContainer.innerHTML = html;
        projetoLeiContainer.classList.remove('hidden');
        gerarPdfBtn.classList.remove('hidden');
        gerarDocxBtn.classList.remove('hidden');

        calcularTotais();
    });

    // util para formatar data (de YYYY-MM-DD para "05 de junho de 2025")
    function formatDataParaAssinatura(dateStr) {
        if (!dateStr) return '';
        // tenta YYYY-MM-DD ou DD/MM/YYYY
        let d, m, y;
        if (dateStr.includes('-')) {
            const parts = dateStr.split('-');
            y = parts[0]; m = parts[1]; d = parts[2];
        } else if (dateStr.includes('/')) {
            const parts = dateStr.split('/');
            d = parts[0]; m = parts[1]; y = parts[2];
        } else return dateStr;
        const meses = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
        const mm = parseInt(m, 10) - 1;
        return `${parseInt(d,10)} de ${meses[mm]} de ${y}`;
    }

    // -----------------------------
    // Geração de DOCX (fidedigno)
    // -----------------------------
    gerarDocxBtn?.addEventListener('click', async () => {
        if (!window.docx) { alert('Biblioteca docx não carregada.'); return; }
        const { Document, Packer, Paragraph, TextRun, HeadingLevel } = window.docx;
        const contentEl = document.getElementById('projeto-lei-gerado');
        if (!contentEl) { alert('Gere o documento antes.'); return; }

        // Extrair informação do preview (já montamos a estrutura antes) -> melhor construir novamente do modelo
        // Para simplificar, recriaremos o conteúdo com os mesmos dados usados no processarBtn
        // Reusar os campos
        const municipio = (document.getElementById('nome-municipio')?.value || '').trim();
        const prefeito = (document.getElementById('nome-prefeito')?.value || '').trim();
        const numeroPL = document.getElementById('numero-pl')?.value || '___/_____';
        const dataDoc = (document.getElementById('data-pl')?.value || '').trim();
        const justificativa = (document.getElementById('justificativa-pl')?.value || '').trim();
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        const secretariaNome = document.getElementById('nome-secretaria')?.value || '';
        const secretariaCargo = document.getElementById('cargo-secretaria')?.value || '';

        // coletar items (mesma lógica que no preview)
        const anulacoes = [];
        anulacaoContainer.querySelectorAll('.anul-item').forEach(div => {
            const sel = div.querySelector('select'); const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const valor = parseNumberFromString(input.value) || 0;
            anulacoes.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });
        const creditos = [];
        creditoContainer.querySelectorAll('.cred-item').forEach(div => {
            const sel = div.querySelector('select'); const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const valor = parseNumberFromString(input.value) || 0;
            creditos.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });

        // criar doc
        const doc = new Document({
            sections: [{
                properties: {},
                children: [
                    new Paragraph({ text: tipo === 'decreto' ? 'DECRETO' : (tipo === 'projetoLei' ? 'PROJETO DE LEI' : 'LEI'), heading: HeadingLevel.HEADING_2, alignment: window.docx.AlignmentType.CENTER }),
                    new Paragraph({ text: '', spacing: { after: 200 } }),
                    new Paragraph({ text: municipio + (dataDoc ? ', ' + dataDoc : ''), alignment: window.docx.AlignmentType.CENTER }),
                    new Paragraph({ text: '' }),
                    // corpo introdutório
                    new Paragraph({ text: tipo === 'decreto' ? `O Prefeito Municipal de ${municipio}, usando de suas atribuições legais, DECRETA:` : (tipo === 'projetoLei' ? `O Prefeito Municipal de ${municipio}, submetendo à apreciação da Câmara Municipal, propõe:` : `O Prefeito Municipal de ${municipio}, faz saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:`) }),
                    new Paragraph({ text: '' })
                ]
            }]
        });

        // inserir artigos (Art. 1º)
        const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
        if (creditos.length > 0) {
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            doc.addSection({
                children: [
                    new Paragraph({ children: [ new TextRun({ text: 'Art. 1º ', bold: true }), new TextRun({ text: `Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender à(s) seguinte(s) dotação(ões):` }) ])
                ]
            });
            // lista de creditos (cada item em parágrafo monoespaçado)
            creditos.forEach(it => {
                doc.addSection({ children: [ new Paragraph({ text: `${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}` }) ] });
            });
            doc.addSection({ children: [ new Paragraph({ text: `TOTAL\nR$ ${formatCurrency(somaCred.toFixed(2))}` }) ] });
        } else {
            doc.addSection({ children: [ new Paragraph({ children: [ new TextRun({ text: 'Art. 1º ', bold: true }), new TextRun({ text: 'Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.' }) ] }) ] });
        }

        // Art.2: fontes
        const somaAnul = anulacoes.reduce((s, it) => s + (it.valor || 0), 0);
        if (anulacoes.length > 0 || parseNumberFromString(valorSuperavitInput?.value) || parseNumberFromString(valorExcessoInput?.value)) {
            const fonteParas = [];
            const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
            const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;
            if (superVal) fonteParas.push(`Superávit Financeiro no valor de R$ ${formatCurrency(superVal.toFixed(2))} (${numeroParaExtensoBR(superVal.toFixed(2))})`);
            if (excessoVal) fonteParas.push(`Excesso de Arrecadação no valor de R$ ${formatCurrency(excessoVal.toFixed(2))} (${numeroParaExtensoBR(excessoVal.toFixed(2))})`);
            anulacoes.forEach(it => fonteParas.push(`Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor.toFixed(2))}`));
            const art2Text = 'Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:';
            doc.addSection({ children: [ new Paragraph({ children: [ new TextRun({ text: 'Art. 2º ', bold: true }), new TextRun({ text: art2Text }) ] }) ] });
            fonteParas.forEach(fp => doc.addSection({ children: [ new Paragraph({ text: fp }) ] }));
            if (anulacoes.length > 0) doc.addSection({ children: [ new Paragraph({ text: `TOTAL (anulação) R$ ${formatCurrency(somaAnul.toFixed(2))}` }) ] });
        }

        // Art.3 e Art.4
        doc.addSection({ children: [ new Paragraph({ children: [ new TextRun({ text: 'Art. 3º ', bold: true }), new TextRun({ text: 'As alterações promovidas passam a fazer parte da LDO e do PPA vigentes.' }) ] }) ] });
        doc.addSection({ children: [ new Paragraph({ children: [ new TextRun({ text: 'Art. 4º ', bold: true }), new TextRun({ text: `Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.` }) ] }) ] });

        // assinatura
        doc.addSection({ children: [
            new Paragraph({ text: '' }),
            new Paragraph({ text: `${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`, alignment: window.docx.AlignmentType.CENTER }),
            new Paragraph({ text: '' }),
            new Paragraph({ text: `${prefeito}`, alignment: window.docx.AlignmentType.CENTER }),
            new Paragraph({ text: 'PREFEITO MUNICIPAL', alignment: window.docx.AlignmentType.CENTER }),
            ...(secretariaNome ? [ new Paragraph({ text: '' }), new Paragraph({ text: `${secretariaNome}`, alignment: window.docx.AlignmentType.CENTER }), new Paragraph({ text: `${secretariaCargo || 'Secretaria'}`, alignment: window.docx.AlignmentType.CENTER }) ] : [])
        ] });

        // justificativa em nova seção/página (se projeto de lei)
        if (document.querySelector('input[name="tipoDocumento"]:checked')?.value === 'projetoLei') {
            doc.addSection({ children: [
                new Paragraph({ text: '', pageBreakBefore: true }),
                new Paragraph({ text: 'JUSTIFICATIVA', heading: HeadingLevel.HEADING_3, alignment: window.docx.AlignmentType.CENTER }),
                new Paragraph({ text: justificativa || '' })
            ] });
        }

        // gerar blob e baixar
        try {
            const blob = await Packer.toBlob(doc);
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = `documento_orcamentario.docx`;
            a.click();
        } catch (err) {
            alert('Erro ao gerar DOCX: ' + err.message);
        }
    });

    // -----------------------------
    // Geração de PDF (jsPDF - texto selecionável)
    // -----------------------------
    gerarPdfBtn?.addEventListener('click', () => {
        const municipio = (document.getElementById('nome-municipio')?.value || '').trim();
        const prefeito = (document.getElementById('nome-prefeito')?.value || '').trim();
        const numeroPL = document.getElementById('numero-pl')?.value || '___/_____';
        const dataDoc = (document.getElementById('data-pl')?.value || '').trim();
        const justificativa = (document.getElementById('justificativa-pl')?.value || '').trim();
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';
        const secretariaNome = document.getElementById('nome-secretaria')?.value || '';
        const secretariaCargo = document.getElementById('cargo-secretaria')?.value || '';

        // coletar items
        const anulacoes = [];
        anulacaoContainer.querySelectorAll('.anul-item').forEach(div => {
            const sel = div.querySelector('select'); const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const valor = parseNumberFromString(input.value) || 0;
            anulacoes.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });
        const creditos = [];
        creditoContainer.querySelectorAll('.cred-item').forEach(div => {
            const sel = div.querySelector('select'); const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)]; const valor = parseNumberFromString(input.value) || 0;
            creditos.push({ codigo: f?.codigo || sel.options[sel.selectedIndex].text, descricao: f?.descricao || '', valor });
        });

        const superVal = parseNumberFromString(valorSuperavitInput?.value) || 0;
        const excessoVal = parseNumberFromString(valorExcessoInput?.value) || 0;

        const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
        const pageWidth = pdf.internal.pageSize.getWidth();
        const margin = 15;
        let y = 20;
        pdf.setFont('Times', 'Normal');
        pdf.setFontSize(12);

        function addTitleCentered(text) {
            pdf.setFontSize(14); pdf.setFont('Times', 'Bold');
            pdf.text(text, pageWidth / 2, y, { align: 'center' });
            y += 8;
            pdf.setFont('Times', 'Normal'); pdf.setFontSize(12);
        }

        function addParagraph(text) {
            const split = pdf.splitTextToSize(text, pageWidth - margin * 2);
            pdf.text(split, margin, y);
            y += split.length * 6;
            if (y > pdf.internal.pageSize.getHeight() - 30) { pdf.addPage(); y = 20; }
        }

        // header
        const titulo = tipo === 'decreto' ? 'DECRETO' : (tipo === 'projetoLei' ? 'PROJETO DE LEI' : 'LEI');
        addTitleCentered(titulo);
        addParagraph(`${municipio}${dataDoc ? ', ' + dataDoc : ''}`);
        if (tipo === 'decreto') addParagraph(`O Prefeito Municipal de ${municipio}, usando de suas atribuições legais,`);
        if (tipo === 'decreto') { addParagraph('DECRETA:'); } else if (tipo === 'projetoLei') addParagraph('O Prefeito Municipal submete ao Legislativo o seguinte projeto:'); else addParagraph('Faço saber que a Câmara Municipal decreta e eu sanciono a seguinte Lei:');

        // Art.1
        if (creditos.length > 0) {
            const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
            const tipoCreditoTxt = (superavitCheck?.checked || excessoCheck?.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            addParagraph(`Art. 1º Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender à(s) seguinte(s) dotação(ões):`);
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
            if (anulacoes.length > 0) {
                const somaAnul = anulacoes.reduce((s, it) => s + (it.valor || 0), 0);
                addParagraph(`TOTAL\nR$ ${formatCurrency(somaAnul.toFixed(2))}`);
            }
        }

        addParagraph('Art. 3º As alterações promovidas passam a integrar a LDO e o PPA vigentes, observadas as disposições legais aplicáveis.');
        addParagraph(`Art. 4º Este ${tipo === 'decreto' ? 'decreto' : (tipo === 'leiFinal' ? 'lei' : 'projeto de lei')} entra em vigor na data de sua publicação.`);

        // assinatura
        y += 10;
        addParagraph(`${municipio}${dataDoc ? ', ' + formatDataParaAssinatura(dataDoc) : ''}`);
        y += 10;
        addParagraph(`${prefeito}`);
        addParagraph('PREFEITO MUNICIPAL');
        if (secretariaNome) {
            y += 10;
            addParagraph(`${secretariaNome}`);
            addParagraph(secretariaCargo || 'Secretaria');
        }

        // justificativa: nova página
        if (document.querySelector('input[name="tipoDocumento"]:checked')?.value === 'projetoLei') {
            pdf.addPage(); y = 20;
            addTitleCentered('JUSTIFICATIVA');
            addParagraph(justificativa || '');
        }

        pdf.save('documento_orcamentario.pdf');
    });

    // recalcula periodicamente (fallback)
    setInterval(calcularTotais, 1000);
});
