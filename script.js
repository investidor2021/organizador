// script.js - Gerador de Projeto/Lei/Decreto com suporte a fichas do Excel, totais e extenso (pt-BR)

// -----------------------------
// Helpers: formatação e extenso
// -----------------------------
function formatCurrency(value) {
    if (value === null || value === undefined || value === '') return '';
    const n = Number(String(value).replace(/\./g, '').replace(',', '.'));
    if (isNaN(n)) return '';
    return n.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// converte número (float ou string) para extenso em reais (ex: "650.320,13" -> "seiscentos e cinquenta mil, trezentos e vinte reais e treze centavos")
// Versão razoavelmente completa para valores até milhões/bilhão — suficiente para uso público municipal.
function numeroParaExtensoBR(valor) {
    if (valor === '' || valor === null || valor === undefined) return '';
    // normalizar
    let n = Number(String(valor).replace(/\./g, '').replace(',', '.'));
    if (isNaN(n)) return '';

    const unidades = ['zero','um','dois','três','quatro','cinco','seis','sete','oito','nove','dez','onze','doze','treze','quatorze','quinze','dezesseis','dezessete','dezoito','dezenove'];
    const dezenas = ['','','vinte','trinta','quarenta','cinquenta','sessenta','setenta','oitenta','noventa'];
    const centenas = ['','cento','duzentos','trezentos','quatrocentos','quinhentos','seiscentos','setecentos','oitocentos','novecentos'];

    function tresDigitosParaExtenso(n) {
        n = Number(n);
        if (n === 0) return '';
        if (n === 100) return 'cem';
        let c = Math.floor(n / 100);
        let d = Math.floor((n % 100) / 10);
        let u = n % 10;
        let texto = '';
        if (n < 20) return unidades[n];
        if (c) texto += centenas[c];
        let resto = n % 100;
        if (resto) {
            if (texto) texto += ' e ';
            if (resto < 20) texto += unidades[resto];
            else {
                let de = Math.floor(resto / 10);
                let un = resto % 10;
                texto += dezenas[de];
                if (un) texto += ' e ' + unidades[un];
            }
        }
        return texto;
    }

    let inteiro = Math.floor(n);
    let centavos = Math.round((n - inteiro) * 100);

    if (inteiro === 0) {
        var resultadoInteiro = 'zero reais';
    } else {
        // separar em grupos de 3 dígitos
        const grupos = [];
        while (inteiro > 0) {
            grupos.push(inteiro % 1000);
            inteiro = Math.floor(inteiro / 1000);
        }
        const nomesGruposSingular = ['', 'mil', 'milhão', 'bilhão', 'trilhão'];
        const nomesGruposPlural = ['', 'mil', 'milhões', 'bilhões', 'trilhões'];

        let partes = [];
        for (let i = grupos.length - 1; i >= 0; i--) {
            const nGrupo = grupos[i];
            if (nGrupo === 0) continue;
            let textoGrupo = tresDigitosParaExtenso(nGrupo);
            let sufixo = '';
            if (i > 0) {
                if (nGrupo === 1) sufixo = ' ' + nomesGruposSingular[i];
                else sufixo = ' ' + nomesGruposPlural[i];
            }
            partes.push((textoGrupo + sufixo).trim());
        }

        resultadoInteiro = partes.join(' e ') + (partsEndsWithPlural(partes) ? ' reais' : ' reais');
    }

    function partsEndsWithPlural(parts) {
        // simplificado: se o número total > 1 -> plural reais
        return true;
    }

    let resultadoFinal = resultadoInteiro;
    if (centavos && centavos > 0) {
        // centavos por extenso
        let cExt = tresDigitosParaExtenso(centavos);
        resultadoFinal += ' e ' + cExt + (centavos === 1 ? ' centavo' : ' centavos');
    }

    return resultadoFinal;
}

// -----------------------------
// Formatação em inputs (ao digitar)
// -----------------------------
function formatCurrencyForInputEvent(e) {
    const el = e.target;
    let digitsOnly = el.value.replace(/\D/g, '');
    if (!digitsOnly) {
        el.value = '';
        return;
    }
    let numberValue = parseInt(digitsOnly, 10) / 100;
    if (isNaN(numberValue)) {
        el.value = '';
        return;
    }
    el.value = numberValue.toLocaleString('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// -----------------------------
// Lógica principal
// -----------------------------
document.addEventListener('DOMContentLoaded', () => {
    // elementos do DOM (garanta que existam no index.html)
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

    const excelFileInput = document.getElementById('excel-file');
    const excelDisplay = document.getElementById('excel-data-display');
    const fichasCount = document.getElementById('fichas-carregadas-count');

    const processarBtn = document.getElementById('processar-btn');
    const gerarPdfBtn = document.getElementById('gerar-pdf-btn');
    const gerarDocxBtn = document.getElementById('gerar-docx-btn');
    const projetoLeiContainer = document.getElementById('projeto-lei-gerado');

    const valorSuperavitInput = document.getElementById('valor-superavit');
    const valorExcessoInput = document.getElementById('valor-excesso');
    const valorInicialSection = document.getElementById('valor-inicial-section');

    let fichasExcel = []; // { codigo, descricao, valor }

    // atualizar visibilidade com controle de habilitação dos botões
    function atualizarVisibilidade() {
        superavitSection?.classList.toggle('hidden', !superavitCheck?.checked);
        excessoSection?.classList.toggle('hidden', !excessoCheck?.checked);
        anulacaoSection?.classList.toggle('hidden', !anulacaoCheck?.checked);

        // creditoSection aparece se qualquer fonte estiver marcada (possível destino)
        const anyFonte = superavitCheck?.checked || excessoCheck?.checked || anulacaoCheck?.checked;
        creditoSection?.classList.toggle('hidden', !anyFonte);

        addAnulacaoBtn.disabled = !anulacaoCheck.checked || fichasExcel.length === 0;
        addCreditoBtn.disabled = !anyFonte || fichasExcel.length === 0;
    }

    // anexar listeners de change
    [superavitCheck, excessoCheck, anulacaoCheck].forEach(chk => {
        if (chk) chk.addEventListener('change', atualizarVisibilidade);
    });

    // Leitura do Excel. Tenta identificar colunas: código (coluna 0), descrição (coluna 1) e valor (coluna que contenha número)
    excelFileInput?.addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function (evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });

            // Identifica cabeçalho (se houver) ou trata a primeira coluna como código
            // Vamos mapear cada linha em { codigo, descricao, valor }
            let rows = sheet.slice(0); // cópia
            // tentativa: se a primeira linha contém texto "Código" ou "Ficha", trata como cabeçalho
            let startIndex = 0;
            const header = rows[0].map(c => String(c).toLowerCase());
            if (header.some(h => /cod|cód|ficha|codigo|descrição|descr/i.test(h))) {
                startIndex = 1;
            }

            fichasExcel = [];
            for (let i = startIndex; i < rows.length; i++) {
                const row = rows[i];
                if (!row || row.length === 0) continue;
                const codigo = (row[0] !== undefined && row[0] !== null) ? String(row[0]).trim() : '';
                const descricao = (row[1] !== undefined && row[1] !== null) ? String(row[1]).trim() : '';
                // tentar encontrar um valor numérico em alguma coluna
                let valor = '';
                for (let c = 2; c < row.length; c++) {
                    const cell = row[c];
                    if (cell === null || cell === '') continue;
                    const cellStr = String(cell).replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
                    if (/^-?\d+(\.\d+)?$/.test(cellStr)) {
                        valor = Number(cellStr);
                        break;
                    }
                }
                // se não encontrou valor nas colunas seguintes, tentar coluna 2 ou 3
                if (valor === '' && row[2] !== undefined && row[2] !== '') {
                    const cellStr = String(row[2]).replace(/\s/g, '').replace(/\./g, '').replace(',', '.');
                    if (/^-?\d+(\.\d+)?$/.test(cellStr)) valor = Number(cellStr);
                }

                // somente linhas com código/descrição
                if (codigo || descricao) {
                    fichasExcel.push({
                        codigo: codigo || descricao,
                        descricao: descricao || codigo,
                        valor: (valor === '' ? null : Number(valor))
                    });
                }
            }

            fichasCount.textContent = `Fichas carregadas: ${fichasExcel.length}`;
            excelDisplay?.classList.remove('hidden');
            atualizarVisibilidade();
        };
        reader.readAsArrayBuffer(file);
    });

    // Função para criar option text legível
    function optionTextFromFicha(f) {
        const valorTxt = f.valor !== null && f.valor !== undefined ? ` - R$ ${formatCurrency(f.valor)}` : '';
        return `${f.codigo} ${f.descricao ? '- ' + f.descricao : ''}${valorTxt}`;
    }

    // Adiciona item de anulação (origem)
    addAnulacaoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const wrapper = document.createElement('div');
        wrapper.className = 'anulacao-item';
        wrapper.style.marginBottom = '8px';

        const select = document.createElement('select');
        select.style.minWidth = '320px';
        fichasExcel.forEach((f, idx) => {
            const opt = document.createElement('option');
            opt.value = idx;
            opt.textContent = optionTextFromFicha(f);
            select.appendChild(opt);
        });

        const valorInput = document.createElement('input');
        valorInput.type = 'text';
        valorInput.placeholder = 'Valor (R$)';
        valorInput.size = 12;
        valorInput.style.marginLeft = '8px';
        valorInput.addEventListener('input', formatCurrencyForInputEvent);

        // se a ficha tiver valor pré-definido, preencher ao selecionar
        select.addEventListener('change', () => {
            const f = fichasExcel[Number(select.value)];
            if (f && f.valor !== null && f.valor !== undefined) {
                valorInput.value = formatCurrency(f.valor);
                calcularTotais();
            }
        });

        // preencher valor inicial com a primeira ficha
        if (fichasExcel[0]) {
            if (fichasExcel[0].valor !== null && fichasExcel[0].valor !== undefined) {
                valorInput.value = formatCurrency(fichasExcel[0].valor);
            }
        }

        const btnRem = document.createElement('button');
        btnRem.type = 'button';
        btnRem.textContent = 'Remover';
        btnRem.style.marginLeft = '8px';
        btnRem.addEventListener('click', () => {
            wrapper.remove();
            calcularTotais();
        });

        wrapper.appendChild(document.createTextNode('Origem: '));
        wrapper.appendChild(select);
        wrapper.appendChild(document.createTextNode(' Valor: '));
        wrapper.appendChild(valorInput);
        wrapper.appendChild(btnRem);

        anulacaoContainer.appendChild(wrapper);
        calcularTotais();
    });

    // Adiciona item de credito (destino)
    addCreditoBtn?.addEventListener('click', () => {
        if (fichasExcel.length === 0) return;
        const wrapper = document.createElement('div');
        wrapper.className = 'credito-item';
        wrapper.style.marginBottom = '8px';

        const select = document.createElement('select');
        select.style.minWidth = '320px';
        fichasExcel.forEach((f, idx) => {
            const opt = document.createElement('option');
            opt.value = idx;
            opt.textContent = optionTextFromFicha(f);
            select.appendChild(opt);
        });

        const valorInput = document.createElement('input');
        valorInput.type = 'text';
        valorInput.placeholder = 'Valor (R$)';
        valorInput.size = 12;
        valorInput.style.marginLeft = '8px';
        valorInput.addEventListener('input', formatCurrencyForInputEvent);

        const btnRem = document.createElement('button');
        btnRem.type = 'button';
        btnRem.textContent = 'Remover';
        btnRem.style.marginLeft = '8px';
        btnRem.addEventListener('click', () => {
            wrapper.remove();
            calcularTotais();
        });

        wrapper.appendChild(document.createTextNode('Destino: '));
        wrapper.appendChild(select);
        wrapper.appendChild(document.createTextNode(' Valor: '));
        wrapper.appendChild(valorInput);
        wrapper.appendChild(btnRem);

        creditoContainer.appendChild(wrapper);
        calcularTotais();
    });

    // calcula totais e atualiza display (pode adicionar campos visuais)
    function calcularTotais() {
        let totalAnul = 0;
        let totalCred = 0;

        anulacaoContainer.querySelectorAll('.anulacao-item').forEach(div => {
            const input = div.querySelector('input');
            if (input && input.value) {
                const n = Number(String(input.value).replace(/\./g, '').replace(',', '.'));
                if (!isNaN(n)) totalAnul += n;
            }
        });

        creditoContainer.querySelectorAll('.credito-item').forEach(div => {
            const input = div.querySelector('input');
            if (input && input.value) {
                const n = Number(String(input.value).replace(/\./g, '').replace(',', '.'));
                if (!isNaN(n)) totalCred += n;
            }
        });

        // mostrar totais — se não existir, cria elementos pequenos abaixo dos containers
        let totAnulEl = document.getElementById('total-anulacao-display');
        let totCredEl = document.getElementById('total-credito-display');

        if (!totAnulEl) {
            totAnulEl = document.createElement('div');
            totAnulEl.id = 'total-anulacao-display';
            totAnulEl.style.marginTop = '8px';
            anulacaoContainer.parentNode.insertBefore(totAnulEl, anulacaoContainer.nextSibling);
        }
        if (!totCredEl) {
            totCredEl = document.createElement('div');
            totCredEl.id = 'total-credito-display';
            totCredEl.style.marginTop = '8px';
            creditoContainer.parentNode.insertBefore(totCredEl, creditoContainer.nextSibling);
        }

        totAnulEl.innerHTML = `<strong>Total Anulação:</strong> R$ ${formatCurrency(totalAnul.toFixed(2))} (${numeroParaExtensoBR(totalAnul.toFixed(2))})`;
        totCredEl.innerHTML = `<strong>Total Crédito:</strong> R$ ${formatCurrency(totalCred.toFixed(2))} (${numeroParaExtensoBR(totalCred.toFixed(2))})`;
    }

    // recalcular totais quando inputs mudam (delegation)
    document.addEventListener('input', (e) => {
        if (e.target && e.target.tagName === 'INPUT' && e.target.type === 'text') {
            // pode ser formato de moeda
            calcularTotais();
        }
    });

    // Geração do documento (HTML final)
    processarBtn?.addEventListener('click', () => {
        const municipio = document.getElementById('nome-municipio')?.value || '';
        const prefeito = document.getElementById('nome-prefeito')?.value || '';
        const numeroPL = document.getElementById('numero-pl')?.value || '___/_____';
        const dataPL = document.getElementById('data-pl')?.value || '';
        const justificativa = document.getElementById('justificativa-pl')?.value || '';
        const tipo = document.querySelector('input[name="tipoDocumento"]:checked')?.value || 'projetoLei';

        // colecionar itens de anulacao e credito
        const anulacoes = [];
        anulacaoContainer.querySelectorAll('.anulacao-item').forEach(div => {
            const sel = div.querySelector('select');
            const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)];
            const valor = input && input.value ? Number(String(input.value).replace(/\./g, '').replace(',', '.')) : null;
            anulacoes.push({
                codigo: f?.codigo || sel.options[sel.selectedIndex].text,
                descricao: f?.descricao || '',
                valor
            });
        });

        const creditos = [];
        creditoContainer.querySelectorAll('.credito-item').forEach(div => {
            const sel = div.querySelector('select');
            const input = div.querySelector('input');
            const f = fichasExcel[Number(sel.value)];
            const valor = input && input.value ? Number(String(input.value).replace(/\./g, '').replace(',', '.')) : null;
            creditos.push({
                codigo: f?.codigo || sel.options[sel.selectedIndex].text,
                descricao: f?.descricao || '',
                valor
            });
        });

        // montar HTML conforme tipo
        let tipoTitulo = tipo === 'projetoLei' ? 'PROJETO DE LEI' : (tipo === 'leiFinal' ? 'LEI' : 'DECRETO');
        let resultadoHTML = `<div id="doc-content" style="font-family: Arial, sans-serif; color:#000; padding:20px;">`;
        resultadoHTML += `<h2 style="text-align:center;">${tipoTitulo}</h2>`;
        resultadoHTML += `<p style="text-align:center;"><strong>${municipio}</strong>${dataPL ? ', ' + dataPL : ''}</p>`;
        resultadoHTML += `<p>O Prefeito Municipal, ${prefeito}, faz saber que:</p>`;

        // Artigos automáticos:
        // Art.1º: abertura do crédito (se houver creditos)
        if (creditos.length > 0) {
            // soma dos creditos
            const somaCred = creditos.reduce((s, it) => s + (it.valor || 0), 0);
            const tipoCreditoTxt = (superavitCheck.checked || excessoCheck.checked) ? 'Crédito Adicional Suplementar' : 'Crédito Adicional Especial';
            resultadoHTML += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir ${tipoCreditoTxt} na importância de R$ ${formatCurrency(somaCred.toFixed(2))} (${numeroParaExtensoBR(somaCred.toFixed(2))}), para atender a(s) seguinte(s) dotação(ões):</p>`;
            resultadoHTML += `<ul>`;
            creditos.forEach(it => {
                resultadoHTML += `<li>${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor ? it.valor.toFixed(2) : 0)}</li>`;
            });
            resultadoHTML += `</ul>`;
            resultadoHTML += `<p><strong>TOTAL</strong><br>R$ ${formatCurrency(somaCred.toFixed(2))}</p>`;
        } else {
            resultadoHTML += `<p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir crédito adicional no orçamento vigente.</p>`;
        }

        // Art.2º: cobertura com anulacoes/excesso/superavit
        if (anulacoes.length > 0 || superavitCheck.checked || excessoCheck.checked) {
            // soma anulacoes
            const somaAnul = anulacoes.reduce((s, it) => s + (it.valor || 0), 0);
            resultadoHTML += `<p><strong>Art. 2º</strong> Para cobertura do crédito autorizado no artigo anterior serão utilizadas as seguintes fontes:</p>`;
            resultadoHTML += `<ul>`;
            if (superavitCheck.checked) {
                const val = Number(String(valorSuperavitInput?.value || '').replace(/\./g, '').replace(',', '.')) || 0;
                resultadoHTML += `<li>Superávit Financeiro no valor de R$ ${formatCurrency(val.toFixed(2))} (${numeroParaExtensoBR(val.toFixed(2))})</li>`;
            }
            if (excessoCheck.checked) {
                const val = Number(String(valorExcessoInput?.value || '').replace(/\./g, '').replace(',', '.')) || 0;
                resultadoHTML += `<li>Excesso de Arrecadação no valor de R$ ${formatCurrency(val.toFixed(2))} (${numeroParaExtensoBR(val.toFixed(2))})</li>`;
            }
            if (anulacoes.length > 0) {
                anulacoes.forEach(it => {
                    resultadoHTML += `<li>Anulação da dotação ${it.codigo} ${it.descricao ? '- ' + it.descricao : ''} – R$ ${formatCurrency(it.valor ? it.valor.toFixed(2) : 0)}</li>`;
                });
                resultadoHTML += `<li><strong>TOTAL (anulação)</strong> R$ ${formatCurrency(somaAnul.toFixed(2))} (${numeroParaExtensoBR(somaAnul.toFixed(2))})</li>`;
            }
            resultadoHTML += `</ul>`;
        }

        // Art.3º: inserir LDO, PPA e legislação citada
        resultadoHTML += `<p><strong>Art. 3º</strong> As alterações promovidas passam a fazer parte da LDO e do PPA vigentes, observadas as disposições legais e normas aplicáveis.</p>`;

        // Art.4º: vigência
        resultadoHTML += `<p><strong>Art. 4º</strong> Este ${tipo === 'decreto' ? 'decreto' : 'projeto de lei' } entra em vigor na data de sua publicação.</p>`;

        resultadoHTML += `</div>`; // fim doc-content

        // JUSTIFICATIVA em página separada para Projeto de Lei
        if (tipo === 'projetoLei') {
            resultadoHTML += `<div style="page-break-before: always;"></div>`;
            resultadoHTML += `<div style="font-family: Arial, sans-serif; padding:20px;">`;
            resultadoHTML += `<h3 style="text-align:center;">JUSTIFICATIVA</h3>`;
            resultadoHTML += `<p>${justificativa.replace(/\n/g, '<br>')}</p>`;
            resultadoHTML += `</div>`;
        }

        projetoLeiContainer.innerHTML = resultadoHTML;
        projetoLeiContainer.classList.remove('hidden');
        gerarPdfBtn.classList.remove('hidden');
        gerarDocxBtn.classList.remove('hidden');

        // recalcula totais ao fim
        calcularTotais();
    });

    // PDF (html2canvas + jspdf)
    gerarPdfBtn?.addEventListener('click', () => {
        const element = document.getElementById('projeto-lei-gerado');
        if (!element) return;
        // ajustar temporariamente estilos para impressão
        html2canvas(element, { scale: 2, useCORS: true }).then(canvas => {
            const imgData = canvas.toDataURL('image/png');
            const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = (canvas.height * pageWidth) / canvas.width;
            pdf.addImage(imgData, 'PNG', 0, 0, pageWidth, pageHeight);
            pdf.save('documento_orcamentario.pdf');
        }).catch(err => {
            alert('Erro ao gerar PDF: ' + err.message);
        });
    });

    // DOCX (usando docx)
    gerarDocxBtn?.addEventListener('click', () => {
        if (!window.docx) {
            alert('Biblioteca docx não está carregada.');
            return;
        }
        const { Document, Packer, Paragraph, TextRun, HeadingLevel } = window.docx;
        const contentEl = document.getElementById('projeto-lei-gerado');
        if (!contentEl) return;

        // construir texto plano com quebras
        // converter nós em texto simples, incluindo listas
        function nodeToText(node) {
            if (!node) return '';
            if (node.nodeType === Node.TEXT_NODE) return node.nodeValue;
            let txt = '';
            node.childNodes.forEach(child => {
                if (child.nodeName === 'BR') txt += '\n';
                else txt += nodeToText(child);
            });
            return txt;
        }
        const fullText = nodeToText(contentEl);

        // transformar em parágrafos simples por quebras de linha
        const paragraphs = fullText.split(/\n{2,}|\r\n{2,}/).map(p => p.trim()).filter(p => p.length > 0);

        const doc = new Document({
            sections: [{
                children: paragraphs.map(p => new Paragraph({ children: [ new TextRun(p) ] }))
            }]
        });

        Packer.toBlob(doc).then(blob => {
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = 'documento_orcamentario.docx';
            a.click();
        }).catch(err => {
            alert('Erro ao gerar DOCX: ' + err.message);
        });
    });

    // recalcular totais a cada mudança
    setInterval(calcularTotais, 1000);
});
