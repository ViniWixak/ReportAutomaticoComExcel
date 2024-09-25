# Relatório Automático de Adesão

Este projeto é uma ferramenta para processar planilhas relacionadas à adesão de colaboradores, atualizando informações com base em dados de pesquisa. Antes de começar, você precisará de Node.js instalado em sua máquina e de duas planilhas: uma de "Pesquisa" contendo os dados relevantes e outra de "Acompanhamento" que serve como template.

## Estrutura das Planilhas

- **Planilha de Pesquisa**: Deve começar com o nome **"Pesquisa"** e conter as seguintes colunas:
  - **A**: Nome Completo (Formato: "Empresa, Nome Colaborador")
  - **B**: Invitees
  - **C**: Respondents
  - **D**: Respondents, %

- **Planilha de Acompanhamento**: Deve começar com o nome **"Acompanhamento"** e conter as seguintes colunas:
  - **A**: Nome do Colaborador
  - **B**: (Coluna vazia)
  - **C**: Invitees (será atualizado)
  - **D**: Respondents (será atualizado)
  - **E**: Respondents, % (será atualizado)

## Instalação

1. Clone o repositório:

   ```
   git clone https://seu-repositorio-url.git
   cd nome-do-repositorio
    ```
2. Instale as dependências:
    ```
    npm install
    ```

## Build do Projeto
Para gerar o executável do projeto, utilize o comando:
    ```
    pkg index.js --output reportAutomatico
    ```

Certifique-se de que o pkg está instalado globalmente. Se não estiver, você pode instalá-lo com o seguinte comando:
    ```
    npm install -g pkg
    ```

## Executando o Programa
Coloque o executável gerado (reportAutomatico) na mesma pasta onde estão as planilhas. Certifique-se de que a planilha de pesquisa comece com "Pesquisa" e a planilha de acompanhamento comece com "Acompanhamento". Execute o programa:
    ```
    ./reportAutomatico
    ```

Após a execução, você encontrará duas planilhas geradas:

- Acompanhamento_Adesao_Atualizado_dd_mm: Com os dados atualizados.
- Acompanhamento Adesão Atualizado dd.mm.aaaa: A planilha de controle atualizada.

## Observações
- O programa deve ser executado no mesmo diretório onde as planilhas estão localizadas.
- A planilha de pesquisa deve ser atualizada periodicamente para garantir que os dados sejam os mais recentes.