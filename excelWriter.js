const ExcelJS = require('exceljs');
const path = require('path');
const { aplicarFormatacaoCondicional } = require('./formatter');

async function escreverDadosPlanilha(dados, nomeBaseArquivo) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Acompanhamento Adesão Atualizado');
    
    // Adicionar os dados à planilha
    worksheet.addRow(dados[0]); // Títulos
    dados.slice(1).forEach(row => worksheet.addRow(row));
    
    // Aplicar a formatação condicional
    aplicarFormatacaoCondicional(worksheet, 5, dados);

    // Nomear o arquivo com a data atual
    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '.');
    const nomeArquivo = `${nomeBaseArquivo} ${dataAtual}.xlsx`;

    // Salvar a planilha
    const caminhoCompleto = path.resolve(__dirname, nomeArquivo);
    await workbook.xlsx.writeFile(caminhoCompleto);

    console.log(`Arquivo salvo com sucesso: ${caminhoCompleto}`);
}

module.exports = { escreverDadosPlanilha };
