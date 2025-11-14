# PROJETO_PLIB_SPAF
projeto em desenvolvido em python de aplicação de um modelo de programação linear inteira binária + ajuste heurístico para a otimização de custos dos projetos do sistema predial de água fria (SPAF)

----------
Instruções de uso:

VERSÃO PYTHON UTILIZADA: 3.13.1

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

todas as planilhas necessárias estão disponiveis na pasta "planilhas" desse repositório.

dentro da pasta "desenho_dxf_rede_figura_1" está disponivel o desenho da rede de exemplo apresentada na figura 1 do tcc no arquivo "desenho rede exemplo figura 1.dxf", caso queira testar o funcionamento do programa, basta inserir o caminho deste arquivo no "caminho_arquivo_dxf" da função main()

Para alterar os valores de velocidade máxima permitidas para os tubos, basta atualizar a planilha "vazões máximas.xlsx", multiplicando o valor da coluna "área (m^2)" pelo valor da vazão que queira adotar, e colar os apenas os valores na coluna 
"vazão (m^3/s)", pois estes serão os novos valores de vazão máxima que atende ao critério de velocidade que o programa vai considerar.

Para alterar os limites de pressão minima exigidos nos pontos de consumo, basta atualizar a coluna "pressão mínima (m.c.a)" da planilha "pesos relativos.xlsx" na linha do respectivo aparelho que queira mudar o critério de pressão minima, por exemplo, para alterar a pressão minima para o aparelho "chuveiro elétrico com registro de pressão", basta atualizar o respectivo valor da coluna "pressão mínima (m.c.a)" na mesma linha do aparelho, que no caso seria a linha 7.

Para alterar os preços adotados para cada componente de "pvc" utilizado na rede, basta atualizar os valores da coluna "preço" da planilha "dados sinapi.xlsx" na linha correspondente do componente que queira atualizar o preço, e caso queira atualizar por valores do sinapi porém de outra data ou estado, a coluna "cod. sinapi" apresenta o código do sinapi para aquele item, e basta buscar por este código dentro da planilha do sinapi, disponivel neste link: https://www.caixa.gov.br/poder-publico/modernizacao-gestao/sinapi/Paginas/default.aspx

Para alterar o valor do limite máximo de velocidade no gráfico do histograma de velocidades, basta atualizar o valor da linha 2509 dentro da função main(), ao lado do texto "velocidade_maxima_do_teste = 3.0", por padrão ela está preenchida como 3 m/s

Para alterar a quantidade de diametros canditados considerados na análise de cada trecho, basta adicionar mais blocos dentro da função "calcular_diametros_adotados", localizada na linha 492, o ajuste é feito da seguinte forma:

                if i+1 < len(registros_vazoes):
                    diametros_nominais.append(registros_vazoes[i+1]["diam_nom"])
                    diametros_internos.append(registros_vazoes[i+1]["diam_interno"])
                    areas.append(registros_vazoes[i+1]["area"])

este trecho do código que fica dentro do "if vazao_reg > flow" deve ser copiado e colado no mesmo alinhamento dos outros "if i+1<len...", e devem ser atualizados os todos os valores (i+numero) para cada novo diametro que for adicionado, para explicar melhor, o trecho acima dentro da função considera apenas um diametro a mais que o minimo.

                if i+1 < len(registros_vazoes):
                    diametros_nominais.append(registros_vazoes[i+1]["diam_nom"])
                    diametros_internos.append(registros_vazoes[i+1]["diam_interno"])
                    areas.append(registros_vazoes[i+1]["area"])
                if i+2 < len(registros_vazoes):
                    diametros_nominais.append(registros_vazoes[i+2]["diam_nom"])
                    diametros_internos.append(registros_vazoes[i+2]["diam_interno"])
                    areas.append(registros_vazoes[i+2]["area"])

esse novo trecho aqui acima considera 2 diametros a mais que o minimo.

depois de ter adicionado os caminhos de todos os arquivos, deve ser iniciado o programa (clicando no "start" do Microsoft Visual Studio 2022), o programa vai executar os calculos deve aparecer o gráfico "distribuição das margens de segurança", que pode ser salvo no computador, e ao fechar a janela do gráfico, deve aparecer a seguinte mensagem:

Gostaria de realizar o orçamento para diâmetros diferentes daqueles escolhidos pela otimização? (responda com sim ou não)

mas antes de responder essa mensagem, é necessário que você copie a tabela apresentada dentro do terminal pelo código, chamada "RESULTADOS DA OTIMIZAÇÃO - DIÂMETROS ESCOLHIDOS E CUSTOS", pois nessa planilha é apresentado os diametros escolhidos para cada diametro pelo processo de otimização, mas também é apresentada a sequencia em que devem ser preenchidos os diametros manuais para a realização do orçamento, minha recomendação é que escreva em uma coluna de uma planilha excel cada diametro que queira utilizar para cada um dos trechos de tubulação, depois de organizar estes dados, volte ao terminal e responda "sim", e será apresentada a seguinte mensagem:

Cole a coluna de diâmetros (um por linha) e pressione Enter duas vezes para finalizar:

então deve ser copiados os dados dos diametros organizados dentro do excel (somente os numeros) e pressionar enter duas vezes, assim serão gerados os gráficos comparativos de comprimento total por diametro e custo total por diametro para a solução otimizada e a solução manual preenchida pelo usuário, ao fechar este gráfico, será apresentado o gráfico de distribuição das velocidades nos trechos da rede, comparando a solução otimizada e a manual.

depois de fechar todos os gráficos, serão apresentados no terminal diversos dados que foram usados pelo programa no processo de resolução do problema.





