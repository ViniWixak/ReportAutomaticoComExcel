const ExcelJS = require('exceljs');

// Função para processar as planilhas
async function processarPlanilhas(caminhoPesquisa, caminhoAcompanhamento) {
    // Carregar a planilha de pesquisa (fonte de dados)
    const workbookPesquisa = new ExcelJS.Workbook();
    await workbookPesquisa.xlsx.readFile(caminhoPesquisa);
    const planilhaPesquisa = workbookPesquisa.worksheets[0];

    // Carregar a planilha de acompanhamento (template)
    const workbookAcompanhamento = new ExcelJS.Workbook();
    await workbookAcompanhamento.xlsx.readFile(caminhoAcompanhamento);
    const planilhaAcompanhamento = workbookAcompanhamento.worksheets[0];

    // Mapeamento dos dados de "Pesquisa de clima"
    const dadosPesquisa = [];
    planilhaPesquisa.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {  // Pula a primeira linha (cabeçalhos)
            const nomeCompleto = row.getCell(1).value; // Coluna A - Nome Completo (Ex: "Heineken Brazil, Nome Colaborador")
            const nomeColaborador = extrairNomeColaborador(nomeCompleto);

            const invitees = row.getCell(2).value; // Coluna B
            const respondents = row.getCell(3).value; // Coluna C
            const respondentsPercent = row.getCell(4).value; // Coluna D

            dadosPesquisa.push({ nomeColaborador, invitees, respondents, respondentsPercent });
        }
    });

    // Atualizar a planilha de acompanhamento
    planilhaAcompanhamento.eachRow((row, rowNumber) => {
        if (rowNumber > 1) {  // Pula a primeira linha (cabeçalhos)
            const nomeColaboradorAcompanhamento = row.getCell(1).value; // Coluna A

            // Encontrar o correspondente na pesquisa
            const dadosCorrespondentes = dadosPesquisa.find(dado => dado.nomeColaborador === nomeColaboradorAcompanhamento);

            if (dadosCorrespondentes) {
                row.getCell(3).value = dadosCorrespondentes.invitees;         // Coluna C - Invitees
                row.getCell(4).value = dadosCorrespondentes.respondents;      // Coluna D - Respondents
                row.getCell(5).value = dadosCorrespondentes.respondentsPercent; // Coluna E - Respondents, %
            }
        }
    });

    // Salvar a planilha de acompanhamento atualizada com data
    const dataAtual = new Date();
    const nomeArquivoSaida = `Acompanhamento_Adesao_Atualizado_${dataAtual.getDate()}_${dataAtual.getMonth() + 1}.xlsx`;

    await workbookAcompanhamento.xlsx.writeFile(nomeArquivoSaida);

    console.log(`Processamento concluído. Arquivo gerado: ${nomeArquivoSaida}`);
}

// Função auxiliar para extrair o nome do colaborador
function extrairNomeColaborador(nomeCompleto) {
    if (nomeCompleto && nomeCompleto.includes(',')) {
        return nomeCompleto.split(',')[1].trim(); // Extrai o nome após a vírgula
    }
    return null; // Retorna null se o nome estiver em um formato inesperado
}

module.exports = {
    processarPlanilhas
};
