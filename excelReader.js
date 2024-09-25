const ExcelJS = require('exceljs');

async function lerDadosPlanilha(nomeArquivo) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(nomeArquivo);
    
    const worksheet = workbook.getWorksheet(1);
    const dados = [];

    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        const rowData = row.values.slice(1); // Ignorar a primeira coluna (id)
        dados.push(rowData);
    });

    return dados;
}

module.exports = { lerDadosPlanilha };
