const xlsx = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');


function getCorPorValor(valor, maximo, minimo) {
    const proporcao = (valor - minimo) / (maximo - minimo);
    
    let red = 0;
    let green = 0;

    if (proporcao <= 0.5) {
        // De vermelho para amarelo
        red = 255; // 100% vermelho
        green = Math.round(255 * (proporcao * 2)); // Aumenta o verde
    } else {
        // De amarelo para verde
        green = 255; // 100% verde
        red = Math.round(255 * (1 - (proporcao - 0.5) * 2)); // Diminui o vermelho
    }

    return { red, green, blue: 0 }; // Azul fixo em 0
}

function extrairNomeColaborador(nomeCompleto) {
    if (typeof nomeCompleto === 'string' && nomeCompleto.includes(',')) {
        const partes = nomeCompleto.split(',');
        if (partes.length > 1) {
            return partes[1].trim();
        }
    }
    return ''; 
}

function encontrarArquivoXls(diretorio, nomePlanilha) {
    const arquivos = fs.readdirSync(diretorio); // Lê os arquivos no diretório
    const arquivo = arquivos.find(arquivo => arquivo.startsWith(nomePlanilha));
    if (arquivo) {
        return path.join(diretorio, arquivo); // Retorna o caminho completo do arquivo
    } else {
        throw new Error(`Arquivo ${nomePlanilha} não encontrado no diretório.`);
    }
}

function obterDataAtual() {
    const hoje = new Date();
    const dia = String(hoje.getDate()).padStart(2, '0'); // Adiciona um zero à esquerda se o dia tiver apenas um dígito
    const mes = String(hoje.getMonth() + 1).padStart(2, '0'); // Meses começam do zero, então adiciona 1
    return `${dia}.${mes}`;
}

const diretorioPlanilhas = './'; 

// Encontra o arquivo da pesquisa de clima
const arquivoPesquisa = encontrarArquivoXls(diretorioPlanilhas, 'Pesquisa');

// Encontra o arquivo de Acompanhamento adesão
const arquivoAcompanhamento = encontrarArquivoXls(diretorioPlanilhas, 'Acompanhamento');

// Lê a planilha "Pesquisa de clima_Adesão dd.MM"
const planilhaPesquisa = xlsx.readFile(arquivoPesquisa);

// Obtenha o intervalo (range) de células da planilha
const abaPesquisa = planilhaPesquisa.Sheets[planilhaPesquisa.SheetNames[1]];
const intervalo = xlsx.utils.decode_range(abaPesquisa['!ref']);  // Obtém o intervalo de células

// Garante que todas as linhas serão lidas, mesmo que estejam vazias
const dadosPesquisa = xlsx.utils.sheet_to_json(abaPesquisa, { header: 1, raw: true, range: intervalo });

// Exibe a quantidade de linhas lidas
console.log(`Linhas lidas da planilha de pesquisa: ${dadosPesquisa.length}`);

// Lê a planilha "Acompanhamento Adesão"
const planilhaAcompanhamento = xlsx.readFile(arquivoAcompanhamento);
const abaAcompanhamento = planilhaAcompanhamento.Sheets[planilhaAcompanhamento.SheetNames[0]];
const dadosAcompanhamento = xlsx.utils.sheet_to_json(abaAcompanhamento, { header: 1 });

// Mapeia os dados de "Pesquisa de clima_Adesão" para atualizar "Acompanhamento Adesão"
for (let i = 2; i < dadosPesquisa.length; i++) {
    const nomeColaboradorPesquisa = extrairNomeColaborador(dadosPesquisa[i][0]); // Extrai o nome do colaborador
    const inviteesPesquisa = dadosPesquisa[i][1];  // Coluna B
    const respondentsPesquisa = dadosPesquisa[i][2];  // Coluna C
    const respondentsPercentPesquisa = dadosPesquisa[i][3];  // Coluna D

    // Busca o nome do colaborador na planilha "Acompanhamento Adesão"
    for (let j = 2; j < dadosAcompanhamento.length; j++) {
        const nomeColaboradorAcompanhamento = dadosAcompanhamento[j][0]; // Coluna A da planilha "Acompanhamento Adesão"

        if (nomeColaboradorPesquisa === nomeColaboradorAcompanhamento) {
            // Atualiza as colunas C, D e E (Invitees, Respondents, Respondents %)
            dadosAcompanhamento[j][2] = inviteesPesquisa;  // Coluna C (Invitees)
            dadosAcompanhamento[j][3] = respondentsPesquisa;  // Coluna D (Respondents)
            dadosAcompanhamento[j][4] = respondentsPercentPesquisa;  // Coluna E (Respondents, %)
            break;  // Para a busca, pois já encontramos o colaborador correspondente
        }
    }
}

// Armazena a primeira linha (títulos)
const titulos = dadosAcompanhamento[0];

// Remove a primeira linha dos dados antes de ordenar
const dadosParaOrdenar = dadosAcompanhamento.slice(1);

dadosParaOrdenar.sort((a, b) => {
    const valorA = String(a[4] || '');
    const valorB = String(b[4] || '');

    // Verifica se algum dos valores contém asterisco
    const contemAsteriscoA = valorA.includes('*');
    const contemAsteriscoB = valorB.includes('*');

    if (contemAsteriscoA && !contemAsteriscoB) {
        return -1; // Coloca valores com asterisco no topo
    }
    if (!contemAsteriscoA && contemAsteriscoB) {
        return 1; // Coloca valores com asterisco no topo
    }
    
    const numeroA = parseFloat(valorA.replace('*', '').trim()) || 0;
    const numeroB = parseFloat(valorB.replace('*', '').trim()) || 0;

    return numeroB - numeroA; // Ordem decrescente dos números
});

const dadosAcompanhamentoOrdenados = [titulos, ...dadosParaOrdenar];

// Criação da nova planilha
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Acompanhamento Adesão');

// Adiciona os títulos
worksheet.addRow(titulos);

// Processa os dados e adiciona as linhas
dadosAcompanhamentoOrdenados.slice(1).forEach(row => {
    worksheet.addRow(row);
});

const valoresColunaE = dadosParaOrdenar.map(row => parseFloat(String(row[4] || '').replace('*', '').trim()) || 0);
const minimo = Math.min(...valoresColunaE);
const maximo = Math.max(...valoresColunaE);

// Converte os dados atualizados de volta para a planilha "Acompanhamento Adesão"
const novosDadosAcompanhamento = xlsx.utils.json_to_sheet(dadosAcompanhamentoOrdenados, { skipHeader: true });

// Aplica a formatação condicional na coluna E
worksheet.getColumn(5).eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    if (rowNumber > 1) { // Ignora o cabeçalho
        const valorE = String(cell.value || '');

        if (valorE.includes('*')) {            
            cell.fill = {
                type: 'pattern',
                pattern: 'none'
            };
        } else {
            const valorNumerico = parseFloat(valorE.replace('*', '').trim()) || 0;

            // Define a cor baseada no valor
            const { red, green } = getCorPorValor(valorNumerico, maximo, minimo);
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: `FF${((1 << 24) + (red << 16) + (green << 8)).toString(16).slice(1).padStart(6, '0')}` }
            };
        }
    }
});

const dataAtual = obterDataAtual();
const nomeArquivoAtualizado = `Acompanhamento Adesão Atualizado ${dataAtual}.xlsx`;

// Salva o arquivo atualizado com a data
workbook.xlsx.writeFile(nomeArquivoAtualizado).then(() => {
    console.log(`Atualização concluída com sucesso! Arquivo salvo como: ${nomeArquivoAtualizado}`);
});