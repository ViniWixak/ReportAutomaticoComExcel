const ExcelReader = require('./excelReader');
const ExcelWriter = require('./excelWriter');
const Formatter = require('./formatter');

async function processarPlanilhas(caminhoPesquisa, caminhoAcompanhamento) {
    // Criar instância do ExcelReader
    const excelReader = new ExcelReader(caminhoPesquisa, caminhoAcompanhamento);
    
    // Ler planilhas
    const { dadosPesquisa, planilhaAcompanhamento } = await excelReader.lerPlanilhas();

    // Criar instância do ExcelWriter
    const excelWriter = new ExcelWriter(planilhaAcompanhamento);
    
    // Escrever dados na planilha de acompanhamento
    await excelWriter.escreverDados(dadosPesquisa);

    // Formatar valores na planilha de acompanhamento
    Formatter.formatarValores(planilhaAcompanhamento);

    // Salvar a planilha de acompanhamento atualizada com data
    const dataAtual = new Date();
    const nomeArquivoSaida = `Acompanhamento_Adesao_Atualizado_${dataAtual.getDate()}_${dataAtual.getMonth() + 1}.xlsx`;
    
    await excelWriter.salvarPlanilha(nomeArquivoSaida);

    console.log(`Processamento concluído. Arquivo gerado: ${nomeArquivoSaida}`);
}

module.exports = {
    processarPlanilhas
};
