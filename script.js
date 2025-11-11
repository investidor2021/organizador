// =====================
// Formatação monetária
// =====================
function formatCurrencyForInput(inputElement) {
    let originalValue = inputElement.value;
    let digitsOnly = originalValue.replace(/\D/g, '');
    if (!digitsOnly) {
        inputElement.value = '';
        return;
    }
    let numberValue = parseInt(digitsOnly, 10) / 100;
    if (isNaN(numberValue)) {
        inputElement.value = '';
        return;
    }
    inputElement.value = numberValue.toLocaleString('pt-BR', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
}

// =====================
// Ação ao carregar página
// =====================
document.addEventListener('DOMContentLoaded', function () {
    const superavitCheck = document.getElementById('fonte_superavit');
    const excessoCheck = document.getElementById('fonte_excesso');
    const anulacaoCheck = document.getElementById('fonte_anulacao');

    const superavitSection = document.getElementById('superavit-valor-section');
    const excessoSection = document.getElementById('excesso-valor-section');
    const anulacaoSection = document.getElementById('anulacao-section');
    const creditoSection = document.getElementById('credito-section');

    const addAnulacaoBtn = document.getElementById('add-anulacao-ficha-btn');
    const addCreditoBtn = document.getElementById('add-credito-ficha-btn');
    const excelFileInput = document.getElementById('excel-file');
    const excelDisplay = document.getElementById('excel-data-display');
    const fichasCount = document.getElementById('fichas-carregadas-count');

    const processarBtn = document.getElementById('processar-btn');
    const gerarPdfBtn = document.getElementById('gerar-pdf-btn');
    const gerarDocxBtn = document.getElementById('gerar-docx-btn');
    const projetoLeiContainer = document.getElementById('projeto-lei-gerado');
    
    // =====================
    // Exibir/ocultar seções
    // =====================
    function atualizarVisibilidade() {
        superavitSection.classList.toggle('hidden', !superavitCheck.checked);
        excessoSection.classList.toggle('hidden', !excessoCheck.checked);
        anulacaoSection.classList.toggle('hidden', !anulacaoCheck.checked);
        creditoSection.classList.toggle('hidden', !(superavitCheck.checked || excessoCheck.checked || anulacaoCheck.checked));

        addAnulacaoBtn.disabled = !anulacaoCheck.checked;
        addCreditoBtn.disabled = !(superavitCheck.checked || excessoCheck.checked || anulacaoCheck.checked);
    }

    superavitCheck.addEventListener('change', atualizarVisibilidade);
    excessoCheck.addEventListener('change', atualizarVisibilidade);
    anulacaoCheck.addEventListener('change', atualizarVisibilidade);

    // =====================
    // Leitura simples do Excel
    // =====================
    excelFileInput.addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function (evt) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const sheet = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
            const total = sheet.length - 1;
            fichasCount.textContent = `Fichas carregadas: ${total}`;
            excelDisplay.classList.remove('hidden');
        };
        reader.readAsArrayBuffer(file);
    });

    // =====================
    // Geração do Projeto de Lei
    // =====================
    processarBtn.addEventListener('click', function () {
        const municipio = document.getElementById('nome-municipio').value;
        const prefeito = document.getElementById('nome-prefeito').value;
        const numeroPL = document.getElementById('numero-pl').value || '___/_____';
        const dataPL = document.getElementById('data-pl').value || '___/___/_____';
        const justificativa = document.getElementById('justificativa-pl').value;

        const tipo = document.querySelector('input[name="tipoDocumento"]:checked').value;
        const tipoTexto = tipo === 'projetoLei' ? 'Projeto de Lei' :
                          tipo === 'leiFinal' ? 'Lei' : 'Decreto';

        const valorSuperavit = document.getElementById('valor-superavit').value || '';
        const valorExcesso = document.getElementById('valor-excesso').value || '';

        let conteudo = `
            <h2 style="text-align:center">${tipoTexto} Nº ${numeroPL}</h2>
            <p style="text-align:center"><strong>${municipio}</strong>, ${dataPL}</p>
            <p>O Prefeito Municipal, ${prefeito}, faz saber que:</p>
            <p><strong>Art. 1º</strong> Fica o Poder Executivo autorizado a abrir crédito adicional ${
                valorSuperavit || valorExcesso ? 'suplementar' : 'especial'
            } no orçamento vigente, conforme valores e dotações especificadas.</p>
            <p><strong>Justificativa:</strong><br>${justificativa}</p>
        `;
        projetoLeiContainer.innerHTML = conteudo;
        projetoLeiContainer.classList.remove('hidden');
        gerarPdfBtn.classList.remove('hidden');
        gerarDocxBtn.classList.remove('hidden');
    });

    // =====================
    // Geração de PDF
    // =====================
    gerarPdfBtn.addEventListener('click', function () {
        const element = document.getElementById('projeto-lei-gerado');
        html2canvas(element).then(canvas => {
            const imgData = canvas.toDataURL('image/png');
            const pdf = new jspdf.jsPDF('p', 'mm', 'a4');
            const pageWidth = pdf.internal.pageSize.getWidth();
            const pageHeight = (canvas.height * pageWidth) / canvas.width;
            pdf.addImage(imgData, 'PNG', 0, 0, pageWidth, pageHeight);
            pdf.save('projeto_de_lei.pdf');
        });
    });

    // =====================
    // Geração de DOCX
    // =====================
    gerarDocxBtn.addEventListener('click', function () {
        const { Document, Packer, Paragraph, TextRun } = window.docx;
        const texto = document.getElementById('projeto-lei-gerado').innerText;
