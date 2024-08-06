function dividirPlanilha() {
    const arquivoInput = document.getElementById('arquivoInput');
    const maxLinhas = parseInt(document.getElementById('maxLinhas').value);

    if (arquivoInput.files.length === 0) {
        alert('Por favor, selecione um arquivo .xlsx primeiro.');
        return;
    }

    const arquivo = arquivoInput.files[0];
    const leitor = new FileReader();

    leitor.onload = function (e) {
        const dados = new Uint8Array(e.target.result);
        const planilhaTrabalho = XLSX.read(dados, { type: 'array' });
        const nomeAba = planilhaTrabalho.SheetNames[0];
        const aba = planilhaTrabalho.Sheets[nomeAba];
        const linhas = XLSX.utils.sheet_to_json(aba, { header: 1 });
        const cabecalho = linhas[0];
        const linhasDados = linhas.slice(1);

        let indiceArquivo = 1;
        for (let i = 0; i < linhasDados.length; i += maxLinhas) {
            const novaAbaDados = [cabecalho, ...linhasDados.slice(i, i + maxLinhas)];
            const novaAba = XLSX.utils.aoa_to_sheet(novaAbaDados);
            const novoLivroTrabalho = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(novoLivroTrabalho, novaAba, `Aba1`);
            XLSX.writeFile(novoLivroTrabalho, `Planilha_Dividida_${indiceArquivo}.xlsx`);
            indiceArquivo++;
        }

        alert('Planilhas geradas com sucesso!');
    };

    leitor.readAsArrayBuffer(arquivo);
}