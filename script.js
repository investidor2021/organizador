function formatCurrencyForInput(inputElement) {
    let originalValue = inputElement.value;

    // 1. Remove todos os caracteres não numéricos.
    //    Isso garante que apenas os dígitos sejam usados para reconstruir o valor.
    let digitsOnly = originalValue.replace(/\D/g, '');

    // 2. Se não houver dígitos, limpa o campo.
    if (!digitsOnly) {
        inputElement.value = '';
        return;
    }

    // 3. Converte a string de dígitos para um número.
    //    Assumimos que os dígitos inseridos representam o valor em centavos.
    //    Por exemplo, "12345" se torna 123.45.
    let numberValue = parseInt(digitsOnly, 10) / 100;

    // 4. Se a conversão falhar (ex: entrada inicial era apenas letras), limpa o campo.
    if (isNaN(numberValue)) {
        inputElement.value = '';
        return;
    }

    // 5. Formata o número para o padrão monetário brasileiro (pt-BR),
    //    que usa '.' como separador de milhar e ',' como separador decimal.
    //    Ex: 1234.56 se torna "1.234,56".
    let formattedValue = numberValue.toLocaleString('pt-BR', {
        minimumFractionDigits: 2, // Sempre mostrar duas casas decimais
        maximumFractionDigits: 2  // No máximo duas casas decimais
    });

    // 6. Atualiza o valor do campo de input com o valor formatado.
    inputElement.value = formattedValue;
}