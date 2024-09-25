const fs = require('fs');
const path = require('path');

// Função para buscar uma planilha pelo prefixo
function getPlanilhaPorPrefixo(prefixo) {
    const arquivos = fs.readdirSync('.'); // Lista todos os arquivos no diretório atual
    const extensoesValidas = ['.xlsx', '.xlsm']; // Extensões de planilhas aceitas

    // Encontra o primeiro arquivo que começa com o prefixo e tem uma das extensões válidas
    const planilha = arquivos.find(arquivo => {
        const ext = path.extname(arquivo).toLowerCase(); // Obtém a extensão do arquivo
        return arquivo.startsWith(prefixo) && extensoesValidas.includes(ext);
    });

    // Se nenhum arquivo for encontrado, lança um erro
    if (!planilha) {
        throw new Error(`Nenhuma planilha encontrada com o prefixo: ${prefixo}`);
    }

    return path.join("./", planilha); // Retorna o caminho completo da planilha encontrada
}

module.exports = {
    getPlanilhaPorPrefixo
};
