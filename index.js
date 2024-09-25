const { lerDadosPlanilha } = require('./excelReader');
const { escreverDadosPlanilha } = require('./excelWriter');
const { processarPlanilhas } = require('./dataProcessor');
const { getPlanilhaPorPrefixo } = require('./fileHelper');

async function executarProcesso() {
    try {
        // Buscar as planilhas que começam com "Pesquisa" e "Acompanhamento"
        const caminhoPlanilhaPesquisa = getPlanilhaPorPrefixo('Pesquisa');
        const caminhoPlanilhaAcompanhamento = getPlanilhaPorPrefixo('Acompanhamento');

        
        const dadosProcessados = processarPlanilhas(caminhoPlanilhaPesquisa, caminhoPlanilhaAcompanhamento);
        
        await escreverDadosPlanilha(dadosProcessados, 'Acompanhamento Adesão Atualizado');
        console.log('Processo concluído com sucesso!');
    } catch (error) {
        console.error('Erro ao executar o processo:', error);
    }
}

executarProcesso();