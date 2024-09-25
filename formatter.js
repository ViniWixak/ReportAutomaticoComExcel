function calcularCorCelula(valor, maximo, minimo) {
    const proporcao = (valor - minimo) / (maximo - minimo);
    
    let red, green;

    if (proporcao <= 0.5) {
        red = 255;
        green = Math.round(255 * (proporcao * 2));
    } else {
        green = 255;
        red = Math.round(255 * (1 - (proporcao - 0.5) * 2));
    }

    return { red, green, blue: 0 };
}

function aplicarFormatacaoCondicional(worksheet, colunaIndex, dados) {
    const valoresColuna = dados.slice(1).map(row => parseFloat(String(row[colunaIndex - 1]).replace('*', '').trim()) || 0);
    const minimo = Math.min(...valoresColuna);
    const maximo = Math.max(...valoresColuna);

    worksheet.getColumn(colunaIndex).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
        if (rowNumber > 1) {
            const valor = String(cell.value || '');

            if (valor.includes('*')) {
                cell.fill = { type: 'pattern', pattern: 'none' };
            } else {
                const valorNumerico = parseFloat(valor.replace('*', '').trim()) || 0;
                const { red, green } = calcularCorCelula(valorNumerico, maximo, minimo);

                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: `FF${((1 << 24) + (red << 16) + (green << 8)).toString(16).slice(1).padStart(6, '0')}` }
                };
            }
        }
    });
}

module.exports = { aplicarFormatacaoCondicional };
