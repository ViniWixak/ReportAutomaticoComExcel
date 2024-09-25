const ExcelJS = require('exceljs');

class ExcelReader {
    constructor(caminhoPesquisa, caminhoAcompanhamento) {
        this.caminhoPesquisa = caminhoPesquisa;
        this.caminhoAcompanhamento = caminhoAcompanhamento;
    }

    async lerPlanilhas() {
        const dadosPesquisa = await this.lerPlanilhaPesquisa();
        const planilhaAcompanhamento = await this.lerPlanilhaAcompanhamento();
        return { dadosPesquisa, planilhaAcompanhamento };
    }

    async lerPlanilhaPesquisa() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.caminhoPesquisa);
        const planilha = workbook.worksheets[0];

        const dadosPesquisa = [];
        planilha.eachRow((row, rowNumber) => {
            if (rowNumber > 1) {
                const nomeCompleto = row.getCell(1).value;
                const nomeColaborador = this.extrairNomeColaborador(nomeCompleto);
                const invitees = row.getCell(2).value;
                const respondents = row.getCell(3).value;
                const respondentsPercent = row.getCell(4).value;

                dadosPesquisa.push({ nomeColaborador, invitees, respondents, respondentsPercent });
            }
        });

        return dadosPesquisa;
    }

    async lerPlanilhaAcompanhamento() {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(this.caminhoAcompanhamento);
        return workbook.worksheets[0];
    }

    extrairNomeColaborador(nomeCompleto) {
        if (nomeCompleto && nomeCompleto.includes(',')) {
            return nomeCompleto.split(',')[1].trim();
        }
        return null;
    }
}

module.exports = ExcelReader;
