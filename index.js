const { processarPlanilhas } = require('./dataProcessor');
const { getPlanilhaPorPrefixo } = require('./fileHelper');

async function executarProcesso() {
    try {
        // Buscar as planilhas que começam com "Pesquisa" e "Acompanhamento"
        const caminhoPlanilhaPesquisa = getPlanilhaPorPrefixo('Pesquisa');
        const caminhoPlanilhaAcompanhamento = getPlanilhaPorPrefixo('Acompanhamento');

        console.log('Iniciando o processamento das planilhas...');
        
        await processarPlanilhas(caminhoPlanilhaPesquisa, caminhoPlanilhaAcompanhamento);
        
        console.log('Processo concluído com sucesso!');
    } catch (error) {
        console.error('Erro ao executar o processo:', error);
    }
}

executarProcesso();