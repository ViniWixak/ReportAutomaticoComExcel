const ExcelJS = require('exceljs');

class Formatter {
    static formatarValores(planilhaAcompanhamento) {
        planilhaAcompanhamento.eachRow((row, rowNumber) => {
            if (rowNumber > 1) { // Ignorar a linha de cabeçalho
                const cellE = row.getCell(5); // Coluna E
                const valueE = cellE.value;

                // Definir cor de fundo para a célula da coluna E
                if (typeof valueE === 'number' && !isNaN(valueE)) {
                    const cor = this.calcularCor(valueE);
                    cellE.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: cor },
                    };
                } else if (typeof valueE === 'string' && valueE.includes('*')) {
                    // Se a célula contém "*", deixar em branco
                    cellE.value = '';
                }

                // Configurar cor de fundo branco para colunas C e D
                row.getCell(3).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFF' }, // Branco
                };
                row.getCell(4).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFF' }, // Branco
                };
            }
        });
    }

    static calcularCor(valor) {
        // Defina a lógica para calcular a cor com base no valor
        const valorMinimo = 0; // Defina o valor mínimo
        const valorMaximo = 100; // Defina o valor máximo
        let percentual = (valor - valorMinimo) / (valorMaximo - valorMinimo);
        percentual = Math.max(0, Math.min(1, percentual)); // Garantir que está entre 0 e 1

        const vermelho = Math.round(255 * (1 - percentual)); // Do vermelho ao verde
        const verde = Math.round(255 * percentual);
        return `${this.converterParaARGB(verde, vermelho, 0)}`;
    }

    static converterParaARGB(verde, vermelho, azul) {
        return `${this.converterParaHex(0)}${this.converterParaHex(vermelho)}${this.converterParaHex(verde)}${this.converterParaHex(azul)}`;
    }

    static converterParaHex(valor) {
        const hex = valor.toString(16).padStart(2, '0').toUpperCase();
        return hex;
    }
}

module.exports = Formatter;
