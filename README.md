# PROJETO_PLIB_SPAF
projeto em desenvolvido em python de aplicação de um modelo de programação linear inteira binária + ajuste heurístico para a otimização de custos dos projetos do sistema predial de água fria (SPAF)

----------
Instruções de uso:

PARA UTILIZAR ESTE PROGRAMA, É NECESSÁRIO QUE SEJA REALIZADO O DOWNLOAD DE TODAS AS PASTAS DESTE REPOSITÓRIO.

DURANTE O DESENVOLVIMENTO DO PROGRAMA, EU UTILIZEI O MICROSOFT VISUAL STUDIO 2022 E RECOMENDO FORTEMENTE QUE UTILIZE ESTE PROGRAMA, POIS NOS TESTES QUE EU FIZ, APARENTEMENTE O SOLVER CBC NÃO FUNCIONA COM O VISUAL STUDIO CODE, PELO MENOS QUANDO EU TENTEI UTILIZAR

Pré-requisitos e Instalação

Para executar este projeto, é necessário ter os seguintes softwares instalados:

Microsoft Visual Studio 2022: Utilizado como ambiente de desenvolvimento.

Solver CBC (COIN-OR Branch-and-Cut): O programa depende deste solver para os cálculos de otimização. As instruções de instalação e o código-fonte estão disponíveis no repositório oficial disponivel no link: https://github.com/coin-or/Cbc

depois de realizar o download das pastas e instalação do solver cbc no computador, o arquivo "Projeto_tcc_visual_studio.py" deve ser aberto no Microsoft Visual Studio 2022, onde deve ser apresentado todo o código do programa. para utilizar o programa, devem ser inseridos dentro da função "main()", entre as linhas 2373 a 2378 do código, os caminhos dos arquivos que foram baixados deste repositório, como segue:

def main():
    caminho_arquivo_dxf = " CAMINHO ARQUIVO DE DESENHO DA REDE .DXF "
    caminho_peso_relativo = " CAMINHO PLANILHA PESOS RELATIVOS .XLSX "
    caminho_vazoes = " CAMINHO PLANILHA VAZÕES MÁXIMAS .XLSX "
    caminho_perda = " CAMINHO PLANILHA PERDA DE CARGA LOCALIZADA .XLSX "
    registros_sinapi = pd.read_excel(" CAMINHO PLANILHA DADOS SINAPI .XLSX ").to_dict(orient="records")
    registros_reducao = carregar_planilha_reducao(" CAMINHO PLANILHA PERDA DE CARGA REDUÇÃO .XLSX ")

todas as planilhas necessárias estão disponiveis na pasta "planilhas" desse repositório



