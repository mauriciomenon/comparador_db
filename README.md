# Comparador

Algoritmo que compara duas tabelas com base nas colunas selecionadas e exibe: <br />
Linhas excluídas da primeira tabela;<br />
Linhas adicionadas à segunda tabela;<br />
Linhas que possuem as colunas selecionadas iguais porem com outros campos diferentes.<br />
<br />
<br />

## VBA
Na pasta 'MACRO_VBA' existe um macro que realiza as comparações diretamente no Excel, o funcionamento é básico e não suporte nenhuma mudança nas tabelas originais utilizadas para teste. <br />

## EXCEL
Na pasta 'EXCEL' existe um algoritmo em python que é capaz de comparar duas tabelas contidas em arquivos Excel (.xlsx), exibindo os resultados em um arquivo Excel que é aberto automaticamente ao final da comparação.<br />
O programa suporta mudar as colunas usadas para a comparação e a ordem das colunas nas tabelas.

## ACCESS

Na pasta 'ACCESS' existe um programa em python que compara as tabelas contidas em arquivos Access (.accdb) ou Excel (.xlsx), exibindo os resultados em tabelas na própria interface.<br />
O programa suporta mudar as colunas usadas para a comparação e a ordem das colunas nas tabelas, também sendo possível utilizar arquivos com múltiplas tabelas.<br />
Não é necessária a instalação de nenhuma das ferramentas da Microsoft para o correto funcionamento do algoritmo, uma vez que é utilizada a ferramenta mdbtools que extrai os dados do arquivo diretamente.<br />
Da maneira que está configurado, o algoritmo utiliza dos executáveis da pasta 'mdbtools' contida nesse repositório, tais arquivos foram compilados para utilização em Windows e não foram testados em outros sistemas operacionais.<br />
O programa é capaz de exportar as tabelas carregadas e o relatório dos resultados para arquivos Excel (.xlsx) utilizando o menu exportar, caso o Excel esteja instalado na máquina o arquivo abre automaticamente. <br />
Para ajuda de como o programa funciona existe um tutorial presente no menu 'ajuda'.

### Bibliotecas necessárias:
`$ pip install subprocess`
`$ pip install pandas`
`$ pip install tkinter`
`$ pip install openpyxl`
`$ pip install pandastable`
