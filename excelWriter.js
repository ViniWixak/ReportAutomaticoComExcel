const ExcelJS = require('exceljs');

class ExcelWriter {
    constructor(planilha) {
        this.planilha = planilha;
    }

    async escreverDados(dados) {
        this.planilha.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const nomeColaboradorAcompanhamento = row.getCell(1).value;
                const dadosCorrespondentes = dados.find(dado => dado.nomeColaborador === nomeColaboradorAcompanhamento);

                if (dadosCorrespondentes) {
                    row.getCell(3).value = dadosCorrespondentes.invitees;
                    row.getCell(4).value = dadosCorrespondentes.respondents;
                    row.getCell(5).value = dadosCorrespondentes.respondentsPercent;
                }
            }
        });
    }

    async salvarPlanilha(nomeArquivo) {
        await this.planilha.workbook.xlsx.writeFile(nomeArquivo);
    }
}

module.exports = ExcelWriter;
