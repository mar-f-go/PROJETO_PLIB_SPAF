import ezdxf
import math
import pandas as pd
import pyomo.environ as pyo
import matplotlib.pyplot as plt
import numpy as np
from collections import defaultdict

def carregar_planilha_pesos_relativos(caminho_peso_relativo):
    """
    Carrega a planilha de pesos relativos
    """
    df = pd.read_excel(caminho_peso_relativo)
    df.rename(columns={
        'aparelho sanitário': 'aparelho_sanitario',
        'peça de utilização': 'peca_de_utilizacao',
        'sigla': 'sigla',
        'vazão de projeto (m^3/s)': 'vazao_de_projeto_m3_s',
        'peso relativo': 'peso_relativo',
        'pressão mínima (m.c.a)': 'pressao_minima_mca'
    }, inplace=True)
    return df.to_dict(orient='records')

def round_coord(coord, decimals=2):
    """
    Arredonda cada valor (x, y, z) da coordenada para o número de casas decimais.
    """
    return (
        round(coord[0], decimals),
        round(coord[1], decimals),
        round(coord[2], decimals)
    )

def ler_textos_e_mtexts(msp):
    """
    Lê entidades TEXT e MTEXT diretamente do Model Space,
    arredondando as coordenadas de inserção.
    """
    resultado = []
    for txt in msp.query('TEXT'):
        conteudo = txt.dxf.text
        coordenadas = round_coord(txt.dxf.insert)
        resultado.append({
            'conteudo': conteudo,
            'coordenadas': coordenadas
        })
    for mtxt in msp.query('MTEXT'):
        conteudo = mtxt.dxf.text
        coordenadas = round_coord(mtxt.dxf.insert)
        resultado.append({
            'conteudo': conteudo,
            'coordenadas': coordenadas
        })
    return resultado

def distancia_3d(inicio, fim):
    return math.sqrt(
        (fim[0] - inicio[0])**2 +
        (fim[1] - inicio[1])**2 +
        (fim[2] - inicio[2])**2
    )

def calcular_angulo_2d(inicio, fim):
    dx = fim[0] - inicio[0]
    dy = fim[1] - inicio[1]
    angulo_rad = math.atan2(dy, dx)
    angulo_graus = math.degrees(angulo_rad)
    return (angulo_graus + 360) % 360

def ler_linhas(msp):
    """
    Lê as propriedades das linhas (coordenadas iniciais e finais),
    arredonda as coordenadas e calcula comprimento e ângulo 2D.
    Adiciona os campos para vazão, diâmetros e comprimentos equivalentes.
    """
    dados_linhas = []
    indice_interno = 1
    for line in msp.query('LINE'):
        inicio = round_coord(line.dxf.start)
        fim = round_coord(line.dxf.end)
        comp = distancia_3d(inicio, fim)
        ang = calcular_angulo_2d(inicio, fim)
        dados_linhas.append({
            'index_interno': indice_interno,
            'inicio': inicio,
            'fim': fim,
            'comprimento': comp,
            'angulo': ang,
            'id': None,
            'textos_associados': [],
            'peso_relativo': None,
            'peso_relativo_total': 0.0,
            'te_saida_lat': 0,
            'te_pass_dir': 0,
            'joelho_45': 0,
            'joelho_90': 0,
            'vazao_m3_s': 0.0,
            "diâmetros nominais adotados:": [],
            "diâmetro interno (m):": [],
            "área (m^2):": [],
            "comprimentos equivalentes": []
        })
        indice_interno += 1
    return dados_linhas

def associar_identificadores(linhas, textos):
    """
    Se o texto for composto apenas por dígitos e sua coordenada coincidir com a
    coordenada final da linha, atribui esse número como ID da linha.
    """
    for texto in textos:
        if texto['conteudo'].isdigit():
            coord_texto = texto['coordenadas']
            identificador = int(texto['conteudo'])
            for linha in linhas:
                if linha['fim'] == coord_texto:
                    linha['id'] = identificador

def associar_textos_nao_numericos(linhas, textos):
    """
    Se o texto não é composto somente por dígitos, verifica se sua coordenada
    coincide com a coordenada inicial (prioridade) ou final da linha e o associa.
    """
    for texto in textos:
        if not texto['conteudo'].isdigit():
            coord_texto = texto['coordenadas']
            linha_encontrada = None
            for ln in linhas:
                if ln['inicio'] == coord_texto:
                    linha_encontrada = ln
                    break
            if linha_encontrada is None:
                for ln in linhas:
                    if ln['fim'] == coord_texto:
                        linha_encontrada = ln
                        break
            if linha_encontrada is not None:
                linha_encontrada['textos_associados'].append(texto['conteudo'])

def associar_peso_relativo(linhas, tabela_pesos_relativos):
    """
    Associa o peso relativo às linhas com base no texto associado (removendo dígitos)
    e comparando com a coluna "sigla" da planilha.
    """
    for ln in linhas:
        for txt in ln['textos_associados']:
            sigla_limpa = ''.join(filter(str.isalpha, txt)).lower()
            if sigla_limpa == 'res':
                continue
            for registro in tabela_pesos_relativos:
                if registro['sigla'].lower() == sigla_limpa:
                    ln['peso_relativo'] = registro['peso_relativo']
                    break

def identificar_tes(linhas):
    """
    Identifica conexões do tipo TE:
      - 2 linhas iniciam na mesma coordenada;
      - 1 linha finaliza nessa mesma coordenada.
    Retorna uma lista de dicionários: { 'coord': (x,y,z), 'line_ids': [...] }.
    """
    inicio_dict = {}
    fim_dict = {}
    for ln in linhas:
        c_inicio = ln['inicio']
        c_fim = ln['fim']
        inicio_dict.setdefault(c_inicio, []).append(ln)
        fim_dict.setdefault(c_fim, []).append(ln)
    lista_tes = []
    for coord, lns in inicio_dict.items():
        if len(lns) == 2 and coord in fim_dict:
            lns_fim = fim_dict[coord]
            if len(lns_fim) == 1:
                te_lines = lns + lns_fim
                line_ids = []
                for l in te_lines:
                    lid = l['id'] if l['id'] is not None else l['index_interno']
                    line_ids.append(lid)
                lista_tes.append({
                    'coord': coord,
                    'line_ids': line_ids
                })
    return lista_tes

def construir_caminhos_siglas_para_res(linhas):
    """
    Constrói o caminho de cada sigla (ex.: 'pt1') até a linha que contenha "res".
    Retorna um dicionário no seguinte formato:
       { sigla: {
             'sigla_completa': <texto>,
             'segmentos_no_caminho': [...],
             'linhas_tes': [],
             'pressao_estatica': 0.0,
             'perda_carga_max_adm': 0.0,
             'msg': ""
           }, ... }
    Cada segmento é um dicionário com as chaves:
         'line_id', 'angle', 'peso_relativo'
    Nesta versão, também incluímos a chave 'coordenadas_iniciais' (obtida do campo "inicio" da linha).
    """
    caminhos = {}
    def eh_reservatorio(ln):
        return any(txt.lower() == "res" for txt in ln['textos_associados'])
    fim_map = {}
    for ln in linhas:
        fim_map[ln['fim']] = ln
    for ln in linhas:
        for txt in ln['textos_associados']:
            sigla_full = txt
            sigla_limpa = ''.join(filter(str.isalpha, txt)).lower()
            if sigla_limpa == "res":
                continue
            if ln['peso_relativo'] is not None:
                segs = []
                pr_inicial = ln['peso_relativo']
                atual = ln
                while True:
                    lid = atual['id'] if atual['id'] is not None else atual['index_interno']
                    # Adiciona também a coordenada inicial da linha ('inicio')
                    segs.append({
                        'line_id': lid,
                        'angle': atual['angulo'],
                        'peso_relativo': pr_inicial,
                        'coordenadas_iniciais': atual['inicio']
                    })
                    if eh_reservatorio(atual):
                        break
                    prox_coord = atual['inicio']
                    if prox_coord in fim_map:
                        atual = fim_map[prox_coord]
                    else:
                        break
                caminhos[sigla_full] = {
                    'sigla_completa': sigla_full,
                    'segmentos_no_caminho': segs,
                    'linhas_tes': [],
                    'pressao_estatica': 0.0,
                    'perda_carga_max_adm': 0.0,
                    'msg': ""
                }
    return caminhos


def somar_pesos_relativos(linhas, caminhos_siglas):
    """
    Para cada sigla, percorre os segmentos do seu caminho e acumula o valor de 'peso_relativo'
    na variável 'peso_relativo_total' na estrutura de dados das linhas – diferenciando as linhas
    que possuem a mesma identificação, mas distintas coordenadas iniciais.
    
    Exemplo: se para a mesma linha id=4 existirem duas entradas com coordenadas (0, 0, 35) e (1, 0, 35),
    cada uma receberá seu próprio somatório.
    """
    # Constrói um dicionário indexado pela chave (linha_id, inicio)
    line_map = {}
    for ln in linhas:
        # Usa ln.get("id") se existir; caso contrário, usa ln["index_interno"]
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        ln["peso_relativo_total"] = 0.0  # inicializa
        line_map[key] = ln

    # Percorre cada sigla e seus segmentos no caminho
    for sigla, info in caminhos_siglas.items():
        segmentos = info.get("segmentos_no_caminho", [])
        for seg in segmentos:
            seg_id = seg.get("line_id")
            # Aqui esperamos que a função construir_caminhos_siglas_para_res tenha armazenado também 
            # as coordenadas iniciais no segmento, na chave 'coordenadas_iniciais'
            coord_ini = seg.get("coordenadas_iniciais")
            key = (seg_id, coord_ini)
            if key in line_map:
                line_map[key]["peso_relativo_total"] += seg.get("peso_relativo", 0.0)

def processar_te_passagem_lateral(linhas, tes_encontradas, caminhos_siglas):
    """
    Para cada TE e para cada sigla, verifica quais segmentos do TE aparecem no caminho da sigla.
    A identificação da linha será feita usando a chave composta (id, coordenadas_iniciais).
    
    Se forem encontrados pelo menos dois segmentos cujos 'line_id' estejam
    presentes em te["line_ids"], eles são ordenados de acordo com
    sua posição no caminho (assumindo que o próprio caminho já esteja ordenado).
    
    Se os ângulos dos dois primeiros segmentos comparados diferirem (mais que uma tolerância), 
    para a primeira linha (linha "anterior") é atribuído o indicador TE saída lateral (te_saida_lat = 1);
    se os ângulos forem iguais, atribuído TE passagem direta (te_pass_dir = 1).
    """
    tol = 1e-6
    # Cria um dicionário para acesso rápido às linhas: chave é (id, inicio)
    line_map = {}
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        line_map[key] = ln

    for te in tes_encontradas:
        # te["line_ids"] contém os ids (números) dos segmentos do TE
        te_ids = te.get("line_ids", [])
        for sigla, info in caminhos_siglas.items():
            segs = info.get("segmentos_no_caminho", [])
            # Filtra os segmentos cujo "line_id" esteja em te_ids
            encontrados = [s for s in segs if s.get("line_id") in te_ids]
            if len(encontrados) < 2:
                continue
            # Ordena os segmentos conforme sua ordem de aparição no caminho
            # Aqui, usamos o índice na lista 'segs' como critério de ordenação.
            encontrados = sorted(encontrados, key=lambda s: segs.index(s))
            primeiro = encontrados[0]
            segundo = encontrados[1]
            diff = abs(primeiro.get("angle", 0) - segundo.get("angle", 0))
            # Utilize a chave composta para buscar a linha correspondente do primeiro segmento
            key_primeiro = (primeiro.get("line_id"), primeiro.get("coordenadas_iniciais"))
            if key_primeiro in line_map:
                linha_primeira = line_map[key_primeiro]
                # Se a diferença de ângulo for 0 (within tolerância), atribuir TE passagem direta; senão, TE saída lateral:
                if abs(diff) < tol:
                    linha_primeira["te_pass_dir"] = 1
                else:
                    linha_primeira["te_saida_lat"] = 1
                # Armazena os dois segmentos do TE no campo 'linhas_tes' da sigla
                info.setdefault("linhas_tes", []).extend([key_primeiro,
                                                          (segundo.get("line_id"), segundo.get("coordenadas_iniciais"))])

def processar_joelhos(linhas, caminhos_siglas):
    """
    Processa os joelhos (45° e 90°) utilizando os ângulos dos segmentos no caminho das siglas.
    Cada linha pode ter, no máximo, 1 indicador entre joelho 45, joelho 90, TE saída lateral ou TE
    passagem direta.
    A identificação da linha é feita via chave composta (id, coordenadas_iniciais).
    """
    angles_90 = {0, 90, 180, 270}
    angles_45 = {45, 135, 225, 315}
    # Cria um dicionário de linhas com chave composta:
    line_map = {}
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        line_map[key] = ln
    
    for sig, data in caminhos_siglas.items():
        segs = data.get("segmentos_no_caminho", [])
        for i in range(len(segs) - 1):
            atual = segs[i]
            prox = segs[i+1]
            a1 = round(atual.get("angle", 0)) % 360
            a2 = round(prox.get("angle", 0)) % 360
            
            # Usa a chave composta para encontrar a linha correspondente ao segmento atual
            key_atual = (atual.get("line_id"), atual.get("coordenadas_iniciais"))
            if key_atual not in line_map:
                continue
            ln_atual = line_map[key_atual]
            
            # Se já existe algum indicador nesta linha, não altera (cada linha deve receber no máximo 1)
            if ln_atual.get("joelho_45", 0) == 1 or ln_atual.get("joelho_90", 0) == 1 or \
               ln_atual.get("te_saida_lat", 0) == 1 or ln_atual.get("te_pass_dir", 0) == 1:
                continue
            
            # Se os ângulos são iguais
            if a1 == a2:
                # MODIFICAÇÃO: Verificação específica para ângulo 0
                if a1 == 0:
                    key_prox = (prox.get("line_id"), prox.get("coordenadas_iniciais"))
                    if key_prox in line_map:
                        ln_prox = line_map[key_prox]
                        try:
                            # Comparar coordenadas X e Y finais da linha próxima com coordenadas X e Y iniciais da linha atual
                            x_inicial_prox = ln_prox.get("inicio")[0]
                            y_inicial_prox = ln_prox.get("inicio")[1]
                            x_final_atual = ln_atual.get("fim")[0]
                            y_final_atual = ln_atual.get("fim")[1]
                            # Se X ou Y forem diferentes, definir joelho_90 = 1
                            if x_inicial_prox != x_final_atual or y_inicial_prox != y_final_atual:
                                ln_atual["joelho_90"] = 1
                        except (IndexError, TypeError):
                            pass  # Se houver erro ao acessar as coordenadas, não faz nada
                continue
            
            # Código para ângulos diferentes
            if (a1 in angles_90 and a2 in angles_90) or (a1 in angles_45 and a2 in angles_45):
                ln_atual["joelho_90"] = 1
            elif a1 in angles_90 and a2 in angles_45:
                ln_atual["joelho_45"] = 1
            elif a1 in angles_45 and a2 in angles_90:
                # Em caso de mudança de angulos entre 45 e 90, mas deveria ser colocado joelho 90, se verifica as coordenadas Z da linha próxima
                key_prox = (prox.get("line_id"), prox.get("coordenadas_iniciais"))
                if key_prox in line_map:
                    ln_prox = line_map[key_prox]
                    try:
                        inicio_z = ln_prox.get("inicio")[2]  # Coordenada Z inicial
                        fim_z = ln_prox.get("fim")[2]        # Coordenada Z final
                        if inicio_z != fim_z:
                            ln_atual["joelho_90"] = 1
                        else:
                            ln_atual["joelho_45"] = 1
                    except (IndexError, TypeError):
                        ln_atual["joelho_45"] = 1  # Valor padrão em caso de erro


def calcular_pressao_estatica_e_perda_carga(linhas, caminhos_siglas, textos, tabela_pesos_relativos):
    """
    Para cada sigla, encontra o texto exato correspondente (ex.: 'cerp1', 'cerp2')
    e utiliza sua coordenada Z para calcular a pressão estática comparada com o Z do texto "res".
    Se a pressão for maior que 40, atribui uma mensagem de alerta.
    
    Utiliza a pressão mínima da tabela de pesos relativos para calcular a perda de carga máxima admissível.
    """
    z_res = None
    for t in textos:
        if t["conteudo"].lower() == "res":
            z_res = t["coordenadas"][2]
            break
    if z_res is None:
        return
    
    # Criar um mapeamento de siglas para pressão mínima
    sigla_pressao_min = {}
    for registro in tabela_pesos_relativos:
        sigla = registro.get('sigla', '').lower()
        pressao_min = registro.get('pressao_minima_mca', 1.0)  # Valor padrão de 1.0 se não encontrado
        sigla_pressao_min[sigla] = pressao_min
            
    for sig_key, info in caminhos_siglas.items():
        z_sigla = None
        for t in textos:
            if t["conteudo"] == sig_key:
                z_sigla = t["coordenadas"][2]
                break
        
        if z_sigla is None:
            info["pressao_estatica"] = 0.0
            info["perda_carga_max_adm"] = -1.0
            info["msg"] = ""
            continue
        
        if z_res > 0 and z_sigla > 0:
            pressao = z_res - z_sigla
        elif z_res < 0 and z_sigla < 0:
            pressao = abs(z_sigla) - abs(z_res)
        elif z_res > 0 and z_sigla < 0:
            pressao = z_res + abs(z_sigla)
        else:
            pressao = abs(z_res) + abs(z_sigla)
        
        info["pressao_estatica"] = pressao
        
        # Extrair a sigla alfanumérica da sigla completa (ex: 'cerp1' -> 'cerp')
        sigla_alfa = ''.join(filter(str.isalpha, sig_key.lower()))
        
        # Obter a pressão mínima para esta sigla
        pressao_min = sigla_pressao_min.get(sigla_alfa, 1.0)  # Valor padrão de 1.0 se não encontrado
        
        # Calcular a perda de carga máxima admissível usando a pressão mínima da tabela
        info["perda_carga_max_adm"] = pressao - pressao_min
        
        if pressao > 40:
            info["msg"] = "a pressão estática neste ponto está superior a 40 m.c.a, não sendo admitido pela norma NBR 5626 2020"
        else:
            info["msg"] = ""


def calcular_vazao(linhas):
    """
    Calcula a vazão m^3/s para cada linha usando a fórmula:
        vazão = 0.0003 * sqrt(peso_relativo_total)
    """
    for ln in linhas:
        if ln["peso_relativo_total"] > 0:
            ln["vazao_m3_s"] = 0.0003 * math.sqrt(ln["peso_relativo_total"])
        else:
            ln["vazao_m3_s"] = 0.0

def carregar_planilha_vazoes_maximas(caminho):
    """
    Carrega a planilha "vazões máximas" contendo as colunas:
      - diâmetro nominal (mm)
      - diâmetro interno (m)
      - área (m^2)
      - vazão (m^3/s)
    """
    df = pd.read_excel(caminho)
    df.rename(columns={
        'diâmetro nominal (mm)': 'diam_nom',
        'diâmetro interno (m)': 'diam_interno',
        'área (m^2)': 'area',
        'vazão (m^3/s)': 'vazao'
    }, inplace=True)
    return df.to_dict(orient='records')

def calcular_diametros_adotados(linhas, registros_vazoes):
    """
    Para cada linha, compara a vazão m^3/s (campo 'vazao_m3_s') com os valores
    da coluna "vazao" da planilha de vazões máximas. Ao encontrar o primeiro 
    registro cuja vazão seja superior à da linha, armazena:
      - Os diâmetros nominais correspondentes e os 3 subsequentes no campo "diâmetros nominais adotados:".
      - Os diâmetros internos correspondentes e os 3 subsequentes no campo "diâmetro interno (m):".
      - As áreas correspondentes e os 3 subsequentes no campo "área (m^2):".
    """
    for ln in linhas:
        flow = ln.get("vazao_m3_s", 0.0)
        diametros_nominais = []
        diametros_internos = []
        areas = []
        for i, reg in enumerate(registros_vazoes):
            vazao_reg = float(reg["vazao"])
            if vazao_reg > flow:
                diametros_nominais.append(reg["diam_nom"])
                diametros_internos.append(reg["diam_interno"])
                areas.append(reg["area"])
                if i+1 < len(registros_vazoes):
                    diametros_nominais.append(registros_vazoes[i+1]["diam_nom"])
                    diametros_internos.append(registros_vazoes[i+1]["diam_interno"])
                    areas.append(registros_vazoes[i+1]["area"])
                if i+2 < len(registros_vazoes):
                    diametros_nominais.append(registros_vazoes[i+2]["diam_nom"])
                    diametros_internos.append(registros_vazoes[i+2]["diam_interno"])
                    areas.append(registros_vazoes[i+2]["area"])
                break
        ln["diâmetros nominais adotados:"] = diametros_nominais
        ln["diâmetro interno (m):"] = diametros_internos
        ln["área (m^2):"] = areas

def carregar_planilha_perda_de_carga(caminho):
    """
    Carrega a planilha "perda de carga localizada" contendo as colunas:
      - diâmetro nominal (mm)
      - joelho 90
      - joelho 45
      - te pass dir
      - te saida lat
      - entrada normal
      - rgl
      - rg
    """
    df = pd.read_excel(caminho)
    df.rename(columns={
        'diâmetro nominal (mm)': 'diam_nom',
        'joelho 90': 'joelho_90',
        'joelho 45': 'joelho_45',
        'te pass dir': 'te_pass_dir',
        'te saida lat': 'te_saida_lat',
        'entrada normal': 'entrada_normal',
        'rgl': 'rgl',
        'rg': 'rg'
    }, inplace=True)
    return df.to_dict(orient='records')

def calcular_comprimentos_equivalentes(linhas, registros_perda):
    """
    Para cada linha, determina o "comprimentos equivalentes" usando os dados da planilha
    "perda de carga localizada". A lógica é:
      - Se a linha possui 1 em algum dos campos "joelho 45", "joelho 90", "te saida lat" ou "te pass dir",
        utiliza a coluna correspondente (prioridade: joelho 45, depois joelho 90, depois te saida lat e te pass dir).
      - Caso contrário, verifica o primeiro texto associado; se for "res", usa "entrada_normal"; se for "rgl", usa "rgl"; se for "rg", usa "rg".
      - Para cada valor em "diâmetros nominais adotados:" da linha, compara com a coluna "diam_nom" da planilha.
        Ao encontrar o primeiro valor igual, obtém o valor correspondente da coluna escolhida e o adiciona ao campo "comprimentos equivalentes".
    """
    for ln in linhas:
        chosen_column = None
        if ln["joelho_45"] == 1:
            chosen_column = "joelho_45"
        elif ln["joelho_90"] == 1:
            chosen_column = "joelho_90"
        elif ln["te_saida_lat"] == 1:
            chosen_column = "te_saida_lat"
        elif ln["te_pass_dir"] == 1:
            chosen_column = "te_pass_dir"
        else:
            if ln["textos_associados"]:
                txt = ln["textos_associados"][0].lower()
                if txt == "res":
                    chosen_column = "entrada_normal"
                elif txt == "rgl":
                    chosen_column = "rgl"
                elif txt == "rg":
                    chosen_column = "rg"
        compr_eq = []
        if chosen_column is not None and "diâmetros nominais adotados:" in ln:
            for diam in ln["diâmetros nominais adotados:"]:
                for reg in registros_perda:
                    try:
                        diam_reg = float(reg["diam_nom"])
                        diam_val = float(diam)
                    except:
                        continue
                    if abs(diam_reg - diam_val) < 1e-6:
                        # Utiliza o método get para evitar KeyError se a coluna não existir
                        valor = reg.get(chosen_column)
                        if valor is not None:
                            compr_eq.append(valor)
                        break
        ln["comprimentos equivalentes"] = compr_eq

def calcular_velocidade_fluido(linhas):
    """
    Para cada linha, percorre os valores do campo "área (m^2):" e calcula a velocidade do fluido.
    Para cada valor, a velocidade é calculada como:
         velocidade = vazao_m3_s / área
    Os resultados são armazenados na lista "velocidade fluido (m/s)" da linha.
    """
    for ln in linhas:
        velocidades = []
        vazao = ln.get("vazao_m3_s", 0.0)
        for area in ln.get("área (m^2):", []):
            try:
                a = float(area)
                if a > 0:
                    velocidades.append(vazao / a)
                else:
                    velocidades.append(0.0)
            except:
                velocidades.append(0.0)
        ln["velocidade fluido (m/s)"] = velocidades

def calcular_reynolds(linhas):
    """
    Para cada linha, calcula Reynolds para cada par correspondente de
    "diâmetro interno (m):" e "velocidade fluido (m/s)" usando a fórmula:
         Reynolds = (velocidade fluido (m/s) * diâmetro interno (m)) / 1e-6
    Os resultados são armazenados como lista no campo "Reynolds" da linha.
    """
    for ln in linhas:
        reynolds_list = []
        diam_interno = ln.get("diâmetro interno (m):", [])
        velocidade  = ln.get("velocidade fluido (m/s)", [])
        for v, d in zip(velocidade, diam_interno):
            try:
                v_val = float(v)
                d_val = float(d)
                reynolds_value = (v_val * d_val) / 1e-6
            except:
                reynolds_value = 0.0
            reynolds_list.append(reynolds_value)
        ln["Reynolds"] = reynolds_list

def calcular_fator_atrito(linhas):
    """
    Calcula o fator de atrito para cada linha.
    Para cada par (diâmetro interno, Reynolds) (listar nos campos "diâmetro interno (m):" e "Reynolds"),
    utiliza a seguinte regra:
      - Se Reynolds > 1e5, então:
            fator de atrito = ( 1 / (-2 * math.log(((e/d)/3.7) + (5.13/(Re**0.89))) ))**0.5
      - Caso contrário, se 5e3 ≤ Reynolds ≤ 1e5 e (e/d) está entre 1e-6 e 1e-2, então:
            fator de atrito = 1.325 / (math.log((e/(3.7*d)) + (5.74/(Re**0.9)))**2)
      - Em qualquer outro caso, o valor será 0.0.
    Nota: A constante e é definida como 0.00006 (rugosidade interna do tubo em m).
    """
    e = 0.00006  # rugosidade interna do tubo (m)
    for ln in linhas:
        fator_list = []
        d_list = ln.get("diâmetro interno (m):", [])
        reynolds_list = ln.get("Reynolds", [])
        for d, Re in zip(d_list, reynolds_list):
            try:
                d_val = float(d)
                Re_val = float(Re)
                if d_val <= 0:
                    fator_list.append(0.0)
                    continue
                rel = e / d_val
                if Re_val > 1e5:
                    ln_term = math.log10(((e / d_val) / 3.7) + (5.13 / (Re_val ** 0.89)))
                    if ln_term == 0:
                        fator_list.append(0.0)
                    else:
                        fator_list.append((1 / (-2 * ln_term)) ** 2)
                elif Re_val >= 5e3 and Re_val <= 1e5 and (rel >= 1e-6 and rel <= 1e-2):
                    ln_term = math.log((e / (3.7 * d_val)) + (5.74 / (Re_val ** 0.9)))
                    if ln_term == 0:
                        fator_list.append(0.0)
                    else:
                        fator_list.append(1.325 / (ln_term ** 2))
                else:
                    fator_list.append(0.0)
            except Exception:
                fator_list.append(0.0)
        ln["fator de atrito"] = fator_list

def calcular_perda_carga_unitaria(linhas):
    """
    Calcula a perda de carga unitária para cada linha usando a fórmula:
       perda de carga unitária = (fator de atrito) * (1 / (diâmetro interno (m))) *
                                 ((velocidade fluido (m/s))^2 / (2 * g))
    Onde a constante g = 9.81 representa a aceleração da gravidade (m/s^2).
    Para cada linha, o cálculo é realizado para cada par correspondente de valores nos campos
    "fator de atrito", "diâmetro interno (m):" e "velocidade fluido (m/s)" e os resultados são armazenados
    como uma lista no campo "perda de carga unitária" da linha.
    """
    g = 9.81  # aceleração da gravidade (m/s^2)
    for ln in linhas:
        perda_list = []
        fator_list = ln.get("fator de atrito", [])
        diam_list  = ln.get("diâmetro interno (m):", [])
        vel_list   = ln.get("velocidade fluido (m/s)", [])
        for f, d, v in zip(fator_list, diam_list, vel_list):
            try:
                f_val = float(f)
                d_val = float(d)
                v_val = float(v)
                if d_val > 0:
                    perda = f_val * (1.0/d_val) * ((v_val**2)/(2*g))
                else:
                    perda = 0.0
            except Exception:
                perda = 0.0
            perda_list.append(perda)
        ln["perda de carga unitária"] = perda_list

def calcular_comprimento_virtual(linhas):
    """
    Calcula o comprimento virtual individualmente para cada item presente em 
    "comprimentos equivalentes" para cada linha, usando a fórmula:
         Comprimento virtual (m) = comprimento + (cada item de "comprimentos equivalentes")
    Arredonda o resultado para 2 casas decimais.
    """
    for ln in linhas:
        comprimento_base = ln.get("comprimento", 0.0)
        eq_list = ln.get("comprimentos equivalentes", [])
        ln["comprimento virtual (m)"] = [round(comprimento_base + eq, 2) for eq in eq_list]

def calcular_perda_carga(linhas):
    """
    Para cada linha, calcula a perda de carga de forma item a item.
    Se o campo "comprimento virtual (m)" não estiver vazio, então para cada item:
         perda de carga = (cada item de "comprimento virtual (m)") * (correspondente perda de carga unitária)
    Caso contrário (se "comprimento virtual (m)" estiver vazio), utiliza o valor do "comprimento" da linha:
         perda de carga = comprimento * (cada item da "perda de carga unitária")
    Os resultados são armazenados no campo "perda de carga" como uma lista.
    """
    for ln in linhas:
        # Tenta obter a lista de comprimento virtual; se estiver vazia, usa o comprimento base da linha.
        comp_virtual = ln.get("comprimento virtual (m)", [])
        if not comp_virtual:
            # Nenhum comprimento virtual disponível: utiliza o valor "comprimento" da linha.
            base = ln.get("comprimento", 0.0)
            perda_unit = ln.get("perda de carga unitária", [])
            perda = [base * pu for pu in perda_unit]
        else:
            # Realiza o cálculo item a item usando os valores fornecidos em "comprimento virtual (m)"
            perda_unit = ln.get("perda de carga unitária", [])
            # Se houver número diferente de itens nas listas, processa até o menor tamanho.
            perda = [cv * pu for cv, pu in zip(comp_virtual, perda_unit)]
        ln["perda de carga"] = perda

def calcular_perda_carga_hidrometro(linhas):
    """
    Para cada linha cujo campo "textos associados" contenha algum texto que inicie com "hidr",
    para cada diâmetro nominal presente no campo "diâmetros nominais adotados:",
    - Seleciona os pares (vazão máxima, diâmetro nominal) da lista definida abaixo cujo diâmetro nominal seja igual.
    - Dentre esses pares, escolhe o primeiro cujo valor de vazão máxima seja maior que o valor de "vazao_m3_s" da linha.
    - Se nenhum par for encontrado para aquele diâmetro nominal, atribui 100 como vazão máxima.
    Em seguida, para cada par selecionado, calcula:
         perda_hid = ((10 * vazao_m3_s) ** 2) / ((vazao_max**2) * 10)
    e armazena esses resultados como uma lista no campo "perda de carga hidrômetro".
    """
    # Lista de pares: (vazão máxima (m^3/s), diâmetro nominal (mm))
    pares = [
        (0.0004166667, 20),
        (0.0008333333, 20),
        (0.0013888889, 20),
        (0.0019444444, 25),
        (0.0027777778, 25),
        (0.0055555556, 40),
        (0.0083333333, 50)
    ]
    for ln in linhas:
        # Verifica se algum texto associado começa com "hidr"
        if not any(txt.lower().startswith("hidr") for txt in ln.get("textos_associados", [])):
            continue
        vazao = ln.get("vazao_m3_s", 0.0)
        di_nom_list = ln.get("diâmetros nominais adotados:", [])
        perda_hidro = []
        # Para cada diâmetro nominal presente na linha...
        for di in di_nom_list:
            try:
                di_val = float(di)
            except:
                continue
            # Selecione os pares cuja diâmetro nominal coincide
            candidatos = [p for p in pares if abs(p[1] - di_val) < 1e-6]
            # Caso haja candidatos, escolha o primeiro cujo vazão máxima seja maior que a vazão da linha
            escolhido = None
            for vazao_max, d_nom in candidatos:
                if vazao_max >= (2*vazao):
                    escolhido = vazao_max
                    break
            # Calcule a perda de carga hidrômetro para este diâmetro:
            try:
                perda = ((10 * vazao) ** 2) / ((escolhido ** 2) * 10)
            # Se não encontrar candidato, atribua 100
            except:
                perda = 100
            perda_hidro.append(perda)
        ln["perda de carga hidrômetro"] = perda_hidro

def atualizar_perda_carga_com_hidrometro(linhas):
    """
    Para cada linha na estrutura de dados (linhas):
      - Se o campo "perda de carga hidrômetro" estiver preenchido (lista não vazia),
        soma elemento a elemento os valores de "perda de carga hidrômetro" com os valores de "perda de carga".
      - Atualiza o campo "perda de carga" com o resultado da soma.
      
    Exemplo:
       Se uma linha tiver:
         "perda de carga": [0.5, 1, 2]
         "perda de carga hidrômetro": [4, 5, 6]
       Então o novo valor será: [0.5+4, 1+5, 2+6] = [4.5, 6, 8].
       
      Caso o campo "perda de carga hidrômetro" esteja vazio, essa linha é ignorada.
    """
    for linha in linhas:
        # Obtenha o valor do campo "perda de carga hidrômetro" (lista)
        hidrometro = linha.get("perda de carga hidrômetro", [])
        if not hidrometro:
            continue  # não há dados para atualizar
        # Obtenha o valor atual do campo "perda de carga" (lista); se não existir, usa lista vazia
        perda = linha.get("perda de carga", [])
        # Realiza a soma elemento a elemento (até o menor tamanho)
        nova_perda = []
        for v1, v2 in zip(perda, hidrometro):
            try:
                nova_perda.append(v1 + v2)
            except Exception:
                nova_perda.append(v1)  # se ocorrer algum problema, mantém o valor original
        # Atualiza o campo "perda de carga" com os novos valores
        linha["perda de carga"] = nova_perda

def buscar_preco_por_diametro(registros_sinapi, diam, tipo, comprimento):
    """
    Procura na tabela 'dados sinapi' uma linha cuja:
      - A coluna "diâmetro nominal entrada" seja igual a diam (com tolerância de 1e-6)
      - E a coluna "tipo" (ignorando maiúsculas/minúsculas) seja igual a tipo.
    Se encontrada:
       - Se o tipo for "tubo", retorna: comprimento * preço (obtido da coluna "preço")
       - Para os demais tipos, retorna apenas o preço encontrado.
    Caso não encontre, retorna 10000.0.
    """
    for reg in registros_sinapi:
        try:
            if abs(float(reg["diâmetro nominal entrada"]) - float(diam)) < 1e-6 and reg["tipo"].strip().lower() == tipo.lower():
                preco = float(reg["preço"])
                if tipo.lower() == "tubo":
                    return comprimento * preco
                else:
                    return preco
        except:
            continue
    return 10000.0

def pertence_ao_te(ln, tes):
    """
    Verifica se a linha ln é o primeiro item (ou seja, o de índice zero) do campo "linhas" em algum TE de tes.
    Retorna True se for, senão False.
    """
    id_val = ln.get("id") if ln.get("id") is not None else ln.get("index_interno")
    for te in tes:
        if te.get("line_ids") and te["line_ids"][0] == id_val:
            return True
    return False

def calcular_preco_diametro(linhas, registros_sinapi, tes):
    """
    Para cada linha, e para cada valor em "diâmetros nominais adotados:",
    realiza os seguintes cálculos:
      - Preço tipo "tubo": Comprimento * (preço obtido para o tipo "tubo").
      - Preço tipo "te": Se a linha pertence a um TE (ou seja, é o primeiro item do TE),
            realiza a busca para o tipo "te"; caso contrário, valor 0.
      - Se "joelho 45" == 1, busca o preço para o tipo "joelho 45".
      - Se "joelho 90" == 1, busca o preço para o tipo "joelho 90".
      - Se em "textos associados" houver item que comece com "hidr", busca o preço para o tipo "hidr".
      - Se em "textos associados" houver "rgl", busca o preço para o tipo "rgl".
      - Se em "textos associados" houver "rg", busca o preço para o tipo "rg".
    Para cada candidato, os valores são reunidos em uma tupla e, em seguida, seu somatório é calculado.
    Os resultados são armazenados nos campos "preço por diâmetro" (lista de tuplas) e "preço total" (lista de somatórios).
    """
    for ln in linhas:
        preco_por_diametro = []
        preco_total = []
        comprimento_base = ln.get("comprimento", 0.0)
        candidatos = ln.get("diâmetros nominais adotados:", [])
        for cand in candidatos:
            comp_prices = []
            # Tipo "tubo"
            price_tubo = buscar_preco_por_diametro(registros_sinapi, cand, "tubo", comprimento_base)
            comp_prices.append(price_tubo)
            # Tipo "te": somente se a linha for o primeiro item de um TE
            if pertence_ao_te(ln, tes):
                price_te = buscar_preco_por_diametro(registros_sinapi, cand, "te", comprimento_base)
            else:
                price_te = 0.0
            comp_prices.append(price_te)
            # Tipo "joelho 45"
            if ln.get("joelho_45", 0) == 1:
                price_jo45 = buscar_preco_por_diametro(registros_sinapi, cand, "joelho 45", comprimento_base)
            else:
                price_jo45 = 0.0
            comp_prices.append(price_jo45)
            # Tipo "joelho 90"
            if ln.get("joelho_90", 0) == 1:
                price_jo90 = buscar_preco_por_diametro(registros_sinapi, cand, "joelho 90", comprimento_base)
            else:
                price_jo90 = 0.0
            comp_prices.append(price_jo90)
            # Tipo "hidr"
            if any(txt.lower().startswith("hidr") for txt in ln.get("textos_associados", [])):
                price_hidr = buscar_preco_por_diametro(registros_sinapi, cand, "hidr", comprimento_base)
            else:
                price_hidr = 0.0
            comp_prices.append(price_hidr)
            # Tipo "rgl"
            if any(txt.lower() == "rgl" for txt in ln.get("textos_associados", [])):
                price_rgl = buscar_preco_por_diametro(registros_sinapi, cand, "rgl", comprimento_base)
            else:
                price_rgl = 0.0
            comp_prices.append(price_rgl)
            # Tipo "rg"
            if any(txt.lower() == "rg" for txt in ln.get("textos_associados", [])):
                price_rg = buscar_preco_por_diametro(registros_sinapi, cand, "rg", comprimento_base)
            else:
                price_rg = 0.0
            comp_prices.append(price_rg)
            preco_por_diametro.append(tuple(comp_prices))
            preco_total.append(sum(comp_prices))
        ln["preço por diâmetro"] = preco_por_diametro
        ln["preço total"] = preco_total

def carregar_planilha_reducao(caminho):
    """
    Carrega a tabela da planilha “perda de carga redução” e retorna uma lista de
    dicionários com as colunas:
        'diâmetro nominal entrada'  →  será renomeada para 'dia_nom_entrada'
        'diâmetro nominal saída'    →  será renomeada para 'dia_nom_saida'
        'coeficiente'
    """
    df = pd.read_excel(caminho)
    df.rename(columns={
        'diâmetro nominal entrada': 'dia_nom_entrada',
        'diâmetro nominal saída': 'dia_nom_saida',
        'coeficiente': 'coeficiente'
    }, inplace=True)
    return df.to_dict(orient="records")

def calcular_perda_carga_reducao(caminhos_siglas, registros_reducao, linhas):
    """
    Calcula a perda de carga por redução para cada sigla, percorrendo o caminho do percurso
    até "res" de trás para frente. Agora, na identificação das linhas, utiliza-se não só o 'id'
    como também as coordenadas iniciais (X, Y, Z) para formar uma chave única.

    Para cada par consecutivo de segmentos do caminho (definindo:
         - linha_anterior: segmento corrente (mais próximo do reservatório ou com 'res')
         - linha_posterior: segmento imediatamente anterior no caminho),
    se o campo "perda de carga redução" da linha posterior estiver vazio, procede com:
         1. Obter das linhas correspondentes (usando a chave (id, inicio)):
              A = lista "diâmetros nominais adotados:" da linha_anterior
              B = lista "diâmetros nominais adotados:" da linha_posterior
              V = lista "velocidade fluido (m/s)" da linha_posterior
         2. Para cada item de A (processado de trás para frente) e para cada item de B (também de trás para frente):
              a) Converte os valores para float (entrada e saida).
              b) Se |entrada – saida| < tol, então perda = 0.0;
              c) Caso contrário, procura em registros_reducao um registro em que:
                     |float(reg["dia_nom_entrada"]) - entrada| < tol  AND  |float(reg["dia_nom_saida"]) - saida| < tol.
                 Se encontrado, extrai o coeficiente e, usando o mesmo índice do item de B na lista V,
                 define: perda = coeficiente * (velocidade²);
                 Se não encontrado, atribui perda = 1000.0.
         3. Armazena, na linha_posterior, a lista de tuplas (entrada, saida, velocidade, perda).
    """
    tol = 1e-6

    # Crie um dicionário para acesso rápido às linhas usando como chave o par (id, inicio)
    line_map = {}
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        coord_iniciais = ln.get("inicio")
        key = (ident, coord_iniciais)
        line_map[key] = ln

    # Para cada sigla, obtenha o caminho
    for sigla, info in caminhos_siglas.items():
        # Tenta primeiro obter o campo "caminho"; se vazio, usa "segmentos_no_caminho"
        caminho = info.get("caminho", [])
        if not caminho:
            caminho = info.get("segmentos_no_caminho", [])
        # Para depuração, exiba os IDs e coordenadas encontrados
        caminho_keys = [ (seg.get("line_id"), seg.get("coordenadas_iniciais")) for seg in caminho if seg.get("line_id") is not None ]
        # Se o caminho não tiver pelo menos dois segmentos, não há par para comparação.
        if len(caminho) < 2:
            continue

        # Percorre o caminho de trás para frente:
        i = len(caminho) - 1  # Último segmento (mais distante do consumo, geralmente com "res")
        while i > 0:
            seg_anterior = caminho[i]    # candidata a linha anterior
            seg_posterior = caminho[i-1]   # candidata a linha posterior

            # Prepare chaves compostas:
            key_anterior = (seg_anterior.get("line_id"), seg_anterior.get("coordenadas_iniciais"))
            key_posterior = (seg_posterior.get("line_id"), seg_posterior.get("coordenadas_iniciais"))
            # Se não encontrar algum dos segmentos na linha, pula o par
            if key_anterior not in line_map or key_posterior not in line_map:
                i -= 1
                continue

            linha_anterior = line_map[key_anterior]
            linha_posterior = line_map[key_posterior]

            # Se linha_posterior já estiver processada, pula
            if linha_posterior.get("perda de carga redução"):
                i -= 1
                continue

            # Obtenha os dados das listas A, B e V
            A = linha_anterior.get("diâmetros nominais adotados:", [])
            B = linha_posterior.get("diâmetros nominais adotados:", [])
            V = linha_posterior.get("velocidade fluido (m/s)", [])
            resultados = []
            if not A or not B or not V:
                linha_posterior["perda de carga redução"] = []
            else:
                # Para cada item de A (linha anterior), iterado de trás para frente
                for idx_a in reversed(range(len(A))):
                    try:
                        entrada = float(A[idx_a])
                    except Exception as e:
                        continue
                    # Para cada item de B (linha posterior), também iterado de trás para frente
                    for idx_b in reversed(range(len(B))):
                        try:
                            saida = float(B[idx_b])
                        except Exception as e:
                            continue
                        if abs(entrada - saida) < tol:
                            perda_val = 0.0
                        else:
                            encontrado = False
                            perda_val = 1000.0  # valor padrão se não for encontrado
                            for reg in registros_reducao:
                                try:
                                    ent_reg = float(reg["dia_nom_entrada"])
                                    sai_reg = float(reg["dia_nom_saida"])
                                except:
                                    continue
                                if abs(ent_reg - entrada) < tol and abs(sai_reg - saida) < tol:
                                    try:
                                        coef = float(reg["coeficiente"])
                                    except:
                                        coef = 0.0
                                    try:
                                        v_val = float(V[idx_b])
                                    except:
                                        v_val = 0.0
                                    perda_val = coef * (v_val ** 2)
                                    encontrado = True
                                    break
                        try:
                            vel = float(V[idx_b])
                        except:
                            vel = 0.0
                        resultados.append((entrada, saida, round(vel,6), round(perda_val,6)))
                linha_posterior["perda de carga redução"] = resultados
            i -= 1


def atualizar_preco_perda_reducao(registros_sinapi, linhas):
    """
    Para cada linha que contenha dados em "perda de carga redução" (lista de tuplas no formato:
    (entrada, saída, velocidade, perda_original)), verifica cada tupla e, se o último valor for diferente de 0
    e diferente de 1000, faz o seguinte:
      - Define 'entrada' = primeiro valor e 'saída' = segundo valor.
      - Procura na tabela sinapi (registros_sinapi) um registro para o qual:
            |float(reg["diâmetro nominal entrada"]) - entrada| < tol  
        e |float(reg["diâmetro nominal saída"]) - saída| < tol.
      - Se for encontrado, extrai o valor de "preço" dessa linha; caso contrário, usa 100.
      - Atualiza a tupla para o formato:
                (entrada, saída, preço, velocidade, perda_original)
    Ao final, atualiza o campo "perda de carga redução" da linha com a lista de tuplas atualizada.
    """
    tol = 1e-6
    for ln in linhas:
        if "perda de carga redução" not in ln or not ln["perda de carga redução"]:
            continue
        novos_resultados = []
        for tup in ln["perda de carga redução"]:
            # Espera-se que "tup" seja uma tupla de 4 elementos: (entrada, saída, velocidade, perda_original)
            try:
                entrada, saida, veloc, perda_original = tup
            except Exception as e:
                continue
            # Se o último valor for 0 ou 1000, não alteramos a tupla
            if abs(perda_original) < tol or abs(perda_original - 1000) < tol:
                novos_resultados.append(tup)
                continue
            # Procura na planilha sinapi um registro que corresponda
            preco = None
            for reg in registros_sinapi:
                try:
                    ent_reg = float(reg["diâmetro nominal entrada"])
                    sai_reg = float(reg["diâmetro nominal saída"])
                except Exception as e:
                    continue
                if abs(ent_reg - entrada) < tol and abs(sai_reg - saida) < tol:
                    try:
                        preco = float(reg["preço"])
                    except:
                        preco = 0.0
                    break
            if preco is None:
                preco = 100
            # Atualiza a tupla para incluir o preço na terceira posição
            novos_resultados.append((entrada, saida, preco, veloc, perda_original))
        ln["perda de carga redução"] = novos_resultados

def filtrar_perda_carga_reducao(linhas):
    """
    Para cada linha, filtra o campo "perda de carga redução" removendo as tuplas cujo último valor seja 0 ou 1000.
    Por exemplo:
       Entrada: [(32.0, 32.0, 0.494244, 0.0), (32.0, 25.0, 0.91, 0.818698, 0.006327), (20.0, 25.0, 0.818698, 1000.0)]
       Saída: [(32.0, 25.0, 0.91, 0.818698, 0.006327)]
    """
    tol = 1e-6
    for linha in linhas:
        resultados = linha.get("perda de carga redução", [])
        # Filtra os resultados: mantém somente aqueles cuja perda (último elemento da tupla) é diferente de zero e de 1000
        filtrados = [tup for tup in resultados if not (abs(tup[-1]) < tol or abs(tup[-1] - 1000) < tol)]
        linha["perda de carga redução"] = filtrados


def solve_linear_model(linhas, caminhos_siglas):
    """
    Monta e resolve o modelo de programação linear para otimizar a escolha de diâmetros,
    com pré-processamento para identificar todas as reduções válidas.
    """
    print("=" * 80)
    print("INICIANDO FUNÇÃO SOLVE_LINEAR_MODEL")
    print("=" * 80)
    
    # Mapeamento entre índices sequenciais e chaves compostas (id, inicio)
    line_map = {}
    idx_to_key = {}
    key_to_idx = {}
    idx = 1
    
    print("\n[1] Criando mapeamento de chaves (id, inicio) para índices simples...")
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        
        if key not in key_to_idx:
            key_to_idx[key] = idx
            idx_to_key[idx] = key
            line_map[key] = ln
            print(f"    Linha: ID={ident}, Início={inicio} => Índice={idx}")
            idx += 1
    
    print(f"\nTotal de trechos mapeados: {len(key_to_idx)}")
    
    # Pré-processamento para identificar todas as reduções válidas
    print("\n[2] Pré-processando reduções válidas entre diâmetros...")
    valid_reductions = {}  # Formato: {(t_anterior, t_posterior): [(j_anterior, j_posterior), ...]}
    tol = 1e-6
    
    # Para cada sigla, processar o caminho
    for sigla, info in caminhos_siglas.items():
        print(f"  Processando sigla: {sigla}")
        path_data = info.get("caminho", [])
        if not path_data:
            path_data = info.get("segmentos_no_caminho", [])
        
        # Converter segmentos para índices
        path_indices = []
        for seg in path_data:
            key = (seg.get("line_id"), seg.get("coordenadas_iniciais"))
            if key in key_to_idx:
                path_indices.append(key_to_idx[key])
        
        # Processar o caminho de trás para frente
        for k in range(len(path_indices)-1, 0, -1):
            t_anterior = path_indices[k]
            t_posterior = path_indices[k-1]
            
            key_anterior = idx_to_key[t_anterior]
            key_posterior = idx_to_key[t_posterior]
            
            linha_anterior = line_map[key_anterior]
            linha_posterior = line_map[key_posterior]
            
            if (t_anterior, t_posterior) not in valid_reductions:
                valid_reductions[(t_anterior, t_posterior)] = []
            
            diams_anteriores = linha_anterior.get("diâmetros nominais adotados:", [])
            diams_posteriores = linha_posterior.get("diâmetros nominais adotados:", [])
            
            for i_ant, diam_ant in enumerate(diams_anteriores):
                for i_post, diam_post in enumerate(diams_posteriores):
                    try:
                        diam_ant_val = float(diam_ant)
                        diam_post_val = float(diam_post)
                        
                        # Diâmetros iguais são válidos (sem redução)
                        if abs(diam_ant_val - diam_post_val) < tol:
                            valid_reductions[(t_anterior, t_posterior)].append((i_ant+1, i_post+1))
                            continue
                        
                        # Para diâmetros diferentes, verificar se existe redução válida
                        for tup in linha_posterior.get("perda de carga redução", []):
                            if len(tup) >= 5 and abs(float(tup[0]) - diam_ant_val) < tol and abs(float(tup[1]) - diam_post_val) < tol:
                                valid_reductions[(t_anterior, t_posterior)].append((i_ant+1, i_post+1))
                                break
                    except (ValueError, TypeError):
                        continue
    
    # Criação do modelo
    model = pyo.ConcreteModel()
    
    # Define o conjunto T usando índices sequenciais
    model.T = pyo.Set(initialize=list(idx_to_key.keys()))
    
    # Define o conjunto TJ para os pares (t,j)
    def TJ_init(model):
        result = []
        for t_idx in model.T:
            key = idx_to_key[t_idx]
            linha = line_map[key]
            diametros = linha.get("diâmetros nominais adotados:", [])
            n = len(diametros)
            for j in range(1, n+1):
                result.append((t_idx, j))
        return result
    
    model.TJ = pyo.Set(initialize=TJ_init)
    
    # Parâmetros e variáveis do modelo
    def D_init(model, t_idx, j):
        key = idx_to_key[t_idx]
        linha = line_map[key]
        try:
            return float(linha["diâmetros nominais adotados:"][j-1])
        except:
            return 0.0
    
    def C_init(model, t_idx, j):
        key = idx_to_key[t_idx]
        linha = line_map[key]
        try:
            return float(linha["preço total"][j-1])
        except:
            return 1000.0
    
    def L_init(model, t_idx, j):
        key = idx_to_key[t_idx]
        linha = line_map[key]
        try:
            return float(linha["perda de carga"][j-1])
        except:
            return 100.0
    
    model.D = pyo.Param(model.TJ, initialize=D_init)
    model.C = pyo.Param(model.TJ, initialize=C_init)
    model.L = pyo.Param(model.TJ, initialize=L_init)
    
    model.x = pyo.Var(model.TJ, domain=pyo.Binary)
    
    def D_choice_rule(model, t_idx):
        return sum(model.D[t_idx, j] * model.x[t_idx, j] for (tt, j) in model.TJ if tt == t_idx)
    
    model.D_choice = pyo.Expression(model.T, rule=D_choice_rule)
    
    # Função objetivo
    def objective_rule(model):
        return sum(model.C[t, j] * model.x[t, j] for (t, j) in model.TJ)
    
    model.obj = pyo.Objective(rule=objective_rule, sense=pyo.minimize)
    
    # Restrições de seleção única e perda de carga
    def selection_rule(model, t_idx):
        return sum(model.x[t_idx, j] for (tt, j) in model.TJ if tt == t_idx) == 1
    
    model.selection = pyo.Constraint(model.T, rule=selection_rule)
    
    # Restrição de perda de carga máxima
    siglas = list(caminhos_siglas.keys())
    model.I = pyo.Set(initialize=siglas)
    
    def head_loss_rule(model, i):
        path_data = caminhos_siglas[i].get("caminho", [])
        if not path_data:
            path_data = caminhos_siglas[i].get("segmentos_no_caminho", [])
        
        path_indices = []
        for seg in path_data:
            key = (seg.get("line_id"), seg.get("coordenadas_iniciais"))
            if key in key_to_idx:
                path_indices.append(key_to_idx[key])
        
        max_head_loss = caminhos_siglas[i].get("perda_carga_max_adm", 100)
        
        return sum(model.L[t_idx, j] * model.x[t_idx, j] 
                  for t_idx in path_indices 
                  for (tt, j) in model.TJ if tt == t_idx) <= max_head_loss
    
    model.head_loss = pyo.Constraint(model.I, rule=head_loss_rule)
    
    # Restrição de monotonicidade dos diâmetros
    model.monotonicity = pyo.ConstraintList()
    
    for i in model.I:
        path_data = caminhos_siglas[i].get("caminho", [])
        if not path_data:
            path_data = caminhos_siglas[i].get("segmentos_no_caminho", [])
        
        path_indices = []
        for seg in path_data:
            key = (seg.get("line_id"), seg.get("coordenadas_iniciais"))
            if key in key_to_idx:
                path_indices.append(key_to_idx[key])
        
        for k in range(len(path_indices)-1, 0, -1):
            t_res_side = path_indices[k]
            t_cons_side = path_indices[k-1]
            
            model.monotonicity.add(model.D_choice[t_res_side] >= model.D_choice[t_cons_side])
    
    # CORREÇÃO: Restrições para combinações inválidas (versão linear)
    print("\n[12] Adicionando restrições para proibir reduções inválidas...")
    model.invalid_constraints = pyo.ConstraintList()
    
    for (t_anterior, t_posterior), valid_pairs in valid_reductions.items():
        if not valid_pairs:
            print(f"    AVISO: Não há reduções válidas para o par {t_anterior}->{t_posterior}")
            continue
            
        print(f"    Restringindo par {t_anterior}->{t_posterior} a {len(valid_pairs)} combinações válidas")
        
        # Obtenha todas as combinações possíveis
        linha_anterior = line_map[idx_to_key[t_anterior]]
        linha_posterior = line_map[idx_to_key[t_posterior]]
        diams_ant = linha_anterior.get("diâmetros nominais adotados:", [])
        diams_post = linha_posterior.get("diâmetros nominais adotados:", [])
        
        # Para cada combinação possível, verifique se é inválida
        for j_ant in range(1, len(diams_ant) + 1):
            for j_post in range(1, len(diams_post) + 1):
                if (j_ant, j_post) not in valid_pairs:
                    # Esta é uma combinação inválida, adicione restrição para proibi-la
                    model.invalid_constraints.add(
                        model.x[t_anterior, j_ant] + model.x[t_posterior, j_post] <= 1
                    )
    
    # Resolver o modelo
    print("\n[13] Resolvendo modelo com solver CBC...")
    solver = pyo.SolverFactory('cbc')
    results = solver.solve(model, tee=True)
    
    # Verificar se o modelo é viável e extrair a solução
    if results.solver.termination_condition == pyo.TerminationCondition.infeasible:
        print("AVISO: O modelo é infeasível - não existe solução que satisfaça todas as restrições.")
        return model, {}
    
    # Extrai a solução
    solution = {}
    for (t_idx, j) in model.TJ:
        try:
            if pyo.value(model.x[t_idx, j]) > 0.5:
                solution[idx_to_key[t_idx]] = j
        except ValueError:
            pass
    
    print(f"\n[14] Solução obtida com {len(solution)} trechos")
    
    # Atualiza custos e perdas de carga devido às reduções na solução final
    update_costs_and_losses_for_solution(solution, linhas, caminhos_siglas, line_map)
    
    return model, solution

def ajustar_diametros_para_margens_negativas(solucao, linhas, caminhos_siglas):
    """
    Ajuste incremental:
      1) Encontra a última redução (mais próxima da sigla).
      2) Remove apenas a redução do trecho alvo.
      3) Escolhe novo diâmetro: d > atual e d ≤ upstream; se d == upstream, aceita sem tupla.
      4) Aplica novas reduções apenas nas fronteiras adjacentes ao trecho alterado (sem duplicar).
      5) Reavalia margem e repete se necessário.
    """
    print("\n" + "=" * 80)
    print("AJUSTANDO DIÂMETROS PARA ELIMINAR MARGENS NEGATIVAS")
    print("=" * 80)

    # map
    line_map = {}
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        line_map[key] = ln

    max_iteracoes = 100
    total_ajustados = 0

    for sigla, info in caminhos_siglas.items():
        caminho = info.get("caminho", []) or info.get("segmentos_no_caminho", [])
        if not caminho:
            continue

        perda_max_adm = info.get("perda_carga_max_adm", 100.0)
        perda_total = calcular_perda_carga_total(caminho, solucao, line_map)
        margem = perda_max_adm - perda_total
        if margem >= 0:
            continue

        print(f"\nSigla {sigla}: Margem inicial = {margem:.6f} m.c.a. (NEGATIVA)")
        total_ajustados += 1

        iteracao = 0
        while iteracao < max_iteracoes:
            iteracao += 1

            # 1) Última redução
            idx_red, key_red, d_upstream = encontrar_ultima_reducao_no_caminho(caminho, solucao, line_map)
            if key_red is None:
                print("  ✗ ERRO: Nenhuma redução encontrada no caminho; não há onde aumentar.")
                break

            j_atual = solucao[key_red]
            ln_cur = line_map[key_red]
            diams = [float(d) for d in ln_cur["diâmetros nominais adotados:"]]
            d_atual = diams[j_atual-1]

            # 2) Escolha do novo diâmetro: d > atual e d ≤ upstream
            prox_j, prox_d = None, None
            for idx, d in enumerate(diams):
                if d <= d_atual:
                    continue
                if d > d_upstream:
                    continue  # nunca criar downstream > upstream
                # caso 1: igualou upstream → sem redução, sempre permitido
                if abs(d - d_upstream) < 1e-6:
                    prox_j, prox_d = idx + 1, d
                    break
                # caso 2: ainda é redução → precisa tupla válida (upstream→d)
                tupla_valida = False
                for tup in ln_cur.get("perda de carga redução", []):
                    if len(tup) >= 5:
                        try:
                            entrada = float(tup[0]); saida = float(tup[1]); perda = float(tup[4])
                            if abs(entrada - d_upstream) < 1e-6 and abs(saida - d) < 1e-6 and \
                               abs(perda) > 1e-6 and abs(perda - 1000) > 1e-6:
                                tupla_valida = True
                                break
                        except:
                            pass
                if tupla_valida:
                    prox_j, prox_d = idx + 1, d
                    break

            if prox_j is None:
                print(f"  ✗ ALERTA: Sem candidato >{d_atual} e ≤{d_upstream} válido; mantendo caminho.")
                break

            print(f"  Iteração {iteracao}: {d_atual}→{prox_d} mm (ID={key_red[0]}, idx={idx_red})")

            # 3) Remover apenas a redução do trecho alvo
            remover_reducao_do_trecho(key_red, solucao, caminho, line_map)

            # 4) Aplicar novo diâmetro
            solucao[key_red] = prox_j

            # 5) Aplicar APENAS novas reduções nas fronteiras adjacentes (sem duplicar)
            aplicar_novas_reducoes_em_torno(caminho, idx_red, solucao, line_map)

            # 6) Reavaliação da margem
            perda_total = calcular_perda_carga_total(caminho, solucao, line_map)
            margem = perda_max_adm - perda_total
            if margem >= 0:
                print(f"  ✓ Margem positiva após {iteracao} iterações: {margem:.6f} m.c.a.")
                break

        if iteracao >= max_iteracoes:
            print("  ✗ ALERTA: Limite de iterações atingido.")

    print(f"\n{'='*80}")
    print(f"TOTAL: {total_ajustados}")
    print(f"{'='*80}\n")
    return solucao

def aplicar_novas_reducoes_em_torno(caminho, idx_alterado, solucao, line_map):
    """
    Aplica APENAS as NOVAS reduções que surgirem nas duas fronteiras adjacentes ao trecho alterado:
    (upstream, alterado) e (alterado, downstream imediato), evitando duplicação via baseline.
    """
    tol = 1e-6

    # Fronteira upstream: (i_up, i_alterado)
    i_up = idx_alterado + 1
    if 0 <= i_up < len(caminho):
        seg_up = caminho[i_up]; seg_dn = caminho[idx_alterado]
        key_up = (seg_up.get("line_id"), seg_up.get("coordenadas_iniciais"))
        key_dn = (seg_dn.get("line_id"), seg_dn.get("coordenadas_iniciais"))
        if key_up in solucao and key_dn in solucao and key_up in line_map and key_dn in line_map:
            j_up = solucao[key_up]; j_dn = solucao[key_dn]
            ln_up = line_map[key_up]; ln_dn = line_map[key_dn]
            try:
                _garantir_baseline(ln_dn)
                d_up = float(ln_up["diâmetros nominais adotados:"][j_up-1])
                d_dn = float(ln_dn["diâmetros nominais adotados:"][j_dn-1])
                if d_up > d_dn + tol:
                    for tup in ln_dn.get("perda de carga redução", []):
                        if len(tup) >= 5:
                            try:
                                entrada = float(tup[0]); saida = float(tup[1])
                                custo = float(tup[2]); perda = float(tup[4])
                                if abs(entrada - d_up) < tol and abs(saida - d_dn) < tol and \
                                   abs(perda) > tol and abs(perda - 1000) > tol:
                                    if not _reducao_ja_aplicada(ln_dn, j_dn, custo, perda):
                                        ln_dn["preço total"][j_dn-1] += custo
                                        ln_dn["perda de carga"][j_dn-1] += perda
                                    break
                            except:
                                pass
                # se d_up == d_dn → sem redução; se d_up < d_dn → inválido, não aplica
            except:
                pass

    # Fronteira downstream: (i_alterado, i_dn2)
    i_dn2 = idx_alterado - 1
    if 0 <= i_dn2 < len(caminho):
        seg_up2 = caminho[idx_alterado]; seg_dn2 = caminho[i_dn2]
        key_up2 = (seg_up2.get("line_id"), seg_up2.get("coordenadas_iniciais"))
        key_dn2 = (seg_dn2.get("line_id"), seg_dn2.get("coordenadas_iniciais"))
        if key_up2 in solucao and key_dn2 in solucao and key_up2 in line_map and key_dn2 in line_map:
            j_up2 = solucao[key_up2]; j_dn2 = solucao[key_dn2]
            ln_up2 = line_map[key_up2]; ln_dn2 = line_map[key_dn2]
            try:
                _garantir_baseline(ln_dn2)
                d_up2 = float(ln_up2["diâmetros nominais adotados:"][j_up2-1])
                d_dn2 = float(ln_dn2["diâmetros nominais adotados:"][j_dn2-1])
                if d_up2 > d_dn2 + tol:
                    for tup in ln_dn2.get("perda de carga redução", []):
                        if len(tup) >= 5:
                            try:
                                entrada = float(tup[0]); saida = float(tup[1])
                                custo = float(tup[2]); perda = float(tup[4])
                                if abs(entrada - d_up2) < tol and abs(saida - d_dn2) < tol and \
                                   abs(perda) > tol and abs(perda - 1000) > tol:
                                    if not _reducao_ja_aplicada(ln_dn2, j_dn2, custo, perda):
                                        ln_dn2["preço total"][j_dn2-1] += custo
                                        ln_dn2["perda de carga"][j_dn2-1] += perda
                                    break
                            except:
                                pass
            except:
                pass

def _reducao_ja_aplicada(ln, j, custo_esp, perda_esp, tol_preco=0.01, tol_perda=1e-6):
    """
    Verifica se a redução (custo_esp, perda_esp) já está aplicada no índice j,
    comparando valores atuais com o baseline original do próprio trecho.
    """
    try:
        _garantir_baseline(ln)
        preco_atual = float(ln["preço total"][j-1]); preco_base = float(ln["preço_total_original"][j-1])
        perda_atual = float(ln["perda de carga"][j-1]); perda_base = float(ln["perda_carga_original"][j-1])
        return (abs((preco_atual - preco_base) - custo_esp) < tol_preco and
                abs((perda_atual - perda_base) - perda_esp) < tol_perda)
    except:
        return False

def remover_reducao_do_trecho(key_trecho, solucao, caminho, line_map):
    """
    Remove APENAS a redução aplicada no trecho key_trecho (par upstream→atual), sem alterar diâmetro.
    """
    tol = 1e-6

    # localizar índice do trecho no caminho
    idx = None
    for i, seg in enumerate(caminho):
        k = (seg.get("line_id"), seg.get("coordenadas_iniciais"))
        if k == key_trecho:
            idx = i
            break
    if idx is None or idx >= len(caminho) - 1:
        return  # não há upstream

    seg_up = caminho[idx + 1]
    key_up = (seg_up.get("line_id"), seg_up.get("coordenadas_iniciais"))
    if key_up not in solucao or key_trecho not in solucao:
        return
    if key_up not in line_map or key_trecho not in line_map:
        return

    j_up = solucao[key_up]
    j_cur = solucao[key_trecho]
    ln_up = line_map[key_up]
    ln_cur = line_map[key_trecho]

    try:
        _garantir_baseline(ln_cur)
        d_up = float(ln_up["diâmetros nominais adotados:"][j_up-1])
        d_cur = float(ln_cur["diâmetros nominais adotados:"][j_cur-1])

        if abs(d_up - d_cur) < tol:
            return  # sem redução aplicada

        for tup in ln_cur.get("perda de carga redução", []):
            if len(tup) >= 5:
                try:
                    entrada = float(tup[0]); saida = float(tup[1])
                    if abs(entrada - d_up) < tol and abs(saida - d_cur) < tol:
                        custo = float(tup[2]); perda = float(tup[4])
                        ln_cur["preço total"][j_cur-1] -= custo
                        ln_cur["perda de carga"][j_cur-1] -= perda
                        return
                except:
                    pass
    except:
        pass


def encontrar_ultima_reducao_no_caminho(caminho, solucao, line_map):
    """
    Retorna (idx_reducao, key_trecho_reduzido, diam_upstream) para a ÚLTIMA redução antes da sigla,
    entendida como a redução mais PRÓXIMA da sigla (primeira encontrada ao varrer sigla→reservatório).
    """
    tol = 1e-6
    for i in range(1, len(caminho)):
        seg_up = caminho[i]       # mais perto do reservatório
        seg_dn = caminho[i-1]     # mais perto da sigla
        key_up = (seg_up.get("line_id"), seg_up.get("coordenadas_iniciais"))
        key_dn = (seg_dn.get("line_id"), seg_dn.get("coordenadas_iniciais"))
        if key_up not in solucao or key_dn not in solucao:
            continue
        if key_up not in line_map or key_dn not in line_map:
            continue
        try:
            j_up = solucao[key_up]; j_dn = solucao[key_dn]
            ln_up = line_map[key_up]; ln_dn = line_map[key_dn]
            d_up = float(ln_up["diâmetros nominais adotados:"][j_up-1])
            d_dn = float(ln_dn["diâmetros nominais adotados:"][j_dn-1])
            if d_up > d_dn + tol:
                return i-1, key_dn, d_up
        except:
            pass
    return None, None, None

def _garantir_baseline(linha):
    """Cria snapshots originais se ainda não existirem (por índice j)."""
    if "preço_total_original" not in linha:
        linha["preço_total_original"] = list(linha["preço total"])
    if "perda_carga_original" not in linha:
        linha["perda_carga_original"] = list(linha["perda de carga"])

def calcular_perda_carga_total(caminho, solucao, line_map):
    """Calcula perda total do caminho."""
    perda_total = 0.0
    for seg in caminho:
        key = (seg.get("line_id"), seg.get("coordenadas_iniciais"))
        if key in solucao and key in line_map:
            j = solucao[key]
            linha = line_map[key]
            try:
                perda = float(linha['perda de carga'][j-1])
                perda_total += perda
            except:
                pass
    return perda_total


def update_costs_and_losses_for_solution(solution, linhas, caminhos_siglas, line_map):
    """
    Recalcula TODAS as reduções de TODOS os caminhos.
    CENÁRIO 2: Detecta e adiciona reduções intermediárias (25→20).
    """
    tol = 1e-6
    atualizacoes = {}
    
    for sigla, info in caminhos_siglas.items():
        path = info.get("caminho", [])
        if not path:
            path = info.get("segmentos_no_caminho", [])
        
        for i in range(len(path) - 1, 0, -1):
            seg_ant = path[i]
            seg_atu = path[i-1]
            key_ant = (seg_ant.get("line_id"), seg_ant.get("coordenadas_iniciais"))
            key_atu = (seg_atu.get("line_id"), seg_atu.get("coordenadas_iniciais"))
            
            if key_ant not in solution or key_atu not in solution:
                continue
            if (key_atu, solution[key_atu]) in atualizacoes:
                continue
            
            try:
                linha_ant = line_map[key_ant]
                linha_atu = line_map[key_atu]
                diam_ant = float(linha_ant["diâmetros nominais adotados:"][solution[key_ant]-1])
                diam_atu = float(linha_atu["diâmetros nominais adotados:"][solution[key_atu]-1])
                
                if abs(diam_ant - diam_atu) < tol:
                    continue
                
                for tup in linha_atu.get("perda de carga redução", []):
                    if len(tup) >= 5:
                        try:
                            if abs(float(tup[0]) - diam_ant) < tol and abs(float(tup[1]) - diam_atu) < tol:
                                linha_atu["preço total"][solution[key_atu]-1] += float(tup[2])
                                linha_atu["perda de carga"][solution[key_atu]-1] += float(tup[4])
                                atualizacoes[(key_atu, solution[key_atu])] = True
                                break
                        except:
                            pass
            except:
                pass

def imprimir_resultados_por_sigla(caminhos_siglas, solucao, linhas):
    """
    Imprime um relatório detalhado para cada sigla, mostrando:
    - Pressão estática e perda de carga máxima admissível
    - Para cada linha no caminho: ID, coordenada inicial, diâmetro escolhido, perda de carga
    - Somatório das perdas de carga de todas as linhas no caminho
    """
    print("\n" + "=" * 100)
    print("RELATÓRIO DETALHADO POR SIGLA".center(100))
    print("=" * 100)
   
    # Cria um mapeamento de linhas por chave composta
    line_map = {}
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        line_map[key] = ln
   
    # Lista para armazenar as margens de segurança calculadas
    margens_de_seguranca = []
   
    # Para cada sigla no sistema
    for sigla, info in caminhos_siglas.items():
        print(f"\nSIGLA: {sigla}")
        print(f"  Pressão estática: {info['pressao_estatica']:.2f} m.c.a.")
        print(f"  Perda de carga máxima admissível: {info['perda_carga_max_adm']:.2f} m.c.a.")
       
        # Alerta se houver mensagem específica
        if info.get('msg', ''):
            print(f"  ALERTA: {info['msg']}")
       
        print("\n  Linhas no caminho:")
        print("  " + "-" * 80)
        print("  {:^8} | {:^30} | {:^10} | {:^15}".format("ID", "Coordenada Inicial", "Diâmetro", "Perda de Carga"))
        print("  " + "-" * 80)
       
        # Obtém o caminho da sigla (segmentos)
        caminho = info.get('caminho', info.get('segmentos_no_caminho', []))
       
        # Calcula a perda de carga total do caminho
        perda_total = 0.0
       
        # Para cada segmento no caminho
        for seg in caminho:
            line_id = seg.get('line_id')
            coord_ini = seg.get('coordenadas_iniciais')
            key = (line_id, coord_ini)
           
            # Formata as coordenadas como string
            coord_str = f"({coord_ini[0]}, {coord_ini[1]}, {coord_ini[2]})"
           
            # Se a linha estiver na solução e no mapeamento
            if key in solucao and key in line_map:
                # Obtem o índice da escolha e a linha correspondente
                j = solucao[key]
                linha = line_map[key]
               
                # Obtem o diâmetro escolhido e a perda de carga correspondente
                try:
                    diametro = linha['diâmetros nominais adotados:'][j-1]
                    perda = float(linha['perda de carga'][j-1])
                    perda_total += perda
                   
                    print("  {:^8} | {:^30} | {:^10} | {:^15.6f}".format(
                        line_id, coord_str, diametro, perda))
                except (IndexError, ValueError) as e:
                    print("  {:^8} | {:^30} | {:^10} | {:^15}".format(
                        line_id, coord_str, "ERRO", "-"))
            else:
                print("  {:^8} | {:^30} | {:^10} | {:^15}".format(
                    line_id, coord_str, "N/A", "-"))
       
        print("  " + "-" * 80)
        print(f"  Perda de carga total do caminho: {perda_total:.6f} m.c.a.")
        
        # Calcula e armazena a margem de segurança
        margem = info['perda_carga_max_adm'] - perda_total
        margens_de_seguranca.append(margem)
        
        print(f"  Margem de segurança: {margem:.6f} m.c.a.")
        print("\n" + "=" * 100)
    
    # Retorna a lista de margens de segurança
    return margens_de_seguranca


def imprimir_resultados_otimizacao(solucao, linhas):
    """
    Imprime os resultados da otimização mostrando o índice, alternativa escolhida,
    diâmetro e preço, finalizando com o somatório total. Os resultados são ordenados pelo ID,
    e para IDs iguais, pela coordenada Z em ordem decrescente.
    """
    print("\n" + "=" * 90)
    print("RESULTADOS DA OTIMIZAÇÃO - DIÂMETROS ESCOLHIDOS E CUSTOS".center(90))
    print("=" * 90)
    
    # Cria um mapeamento de linhas por chave composta
    line_map = {}
    key_to_idx = {}  # Mapeamento de chaves para índices
    idx = 1
    
    for ln in linhas:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        line_map[key] = ln
        key_to_idx[key] = idx
        idx += 1
    
    # Calcula o custo total e imprime os detalhes de cada trecho
    custo_total = 0.0
    
    print(f"{'ID':<6} | {'Índice':<8} | {'Coordenadas Iniciais':<30} | {'Alt.':<4} | {'Diâmetro':<8} | {'Preço (R$)':<12}")
    print("-" * 90)
    
    # Ordenar solução pelo ID e depois pela coordenada Z (decrescente)
    sorted_items = []
    for t, j in solucao.items():
        ident, coord_ini = t  # t é uma tupla (id, coordenadas)
        idx_linha = key_to_idx.get(t, 9999)
        
        # Extrair coordenada Z (se disponível)
        z_coord = coord_ini[2] if coord_ini and len(coord_ini) > 2 else -float('inf')
        
        sorted_items.append((ident, z_coord, t, j, idx_linha))
    
    # Ordenar por ID (primeiro elemento) e depois por Z decrescente (negativo da segunda coordenada)
    sorted_items.sort(key=lambda x: (x[0] is None, x[0], -x[1]))  # None vai para o final, Z em ordem decrescente
    
    for ident, z_coord, t, j, idx_linha in sorted_items:
        try:
            linha = line_map.get(t)
            if linha is None:
                print(f"{'AVISO':<6} | {'N/A':<8} | Linha não encontrada para trecho {t}")
                continue
            
            # Verifica se ident é None e usa um valor padrão
            ident_str = str(ident) if ident is not None else "N/A"
            
            # Verifica se coord_ini e suas componentes são válidas
            coord_ini = t[1]
            if coord_ini is None or not all(isinstance(c, (int, float)) for c in coord_ini):
                coord_str = "Coordenadas inválidas"
            else:
                # Formata a coordenada como string de forma segura
                coord_str = f"({coord_ini[0]:.1f}, {coord_ini[1]:.1f}, {coord_ini[2]:.1f})"
            
            # Verifica j e acessa os arrays com segurança
            if j is None or not isinstance(j, int):
                j_str = "N/A"
                diametro_str = "N/A"
                preco_str = "N/A"
            else:
                j_str = str(j)
                
                # Acessa diâmetro e preço com segurança
                diametros = linha.get('diâmetros nominais adotados:', [])
                precos = linha.get('preço total', [])
                
                if j <= 0 or j > len(diametros):
                    diametro_str = "N/A"
                else:
                    diametro = diametros[j-1]
                    diametro_str = str(diametro) if diametro is not None else "N/A"
                
                if j <= 0 or j > len(precos):
                    preco_str = "N/A"
                else:
                    preco = precos[j-1]
                    if preco is not None and isinstance(preco, (int, float)):
                        preco_str = f"R$ {preco:.2f}"
                        custo_total += preco
                    else:
                        preco_str = "N/A"
            
            # Imprime linha com formatação segura, incluindo o índice
            print(f"{ident_str:<6} | {idx_linha:<8} | {coord_str:<30} | {j_str:<4} | {diametro_str:<8} | {preco_str:<12}")
            
        except Exception as e:
            print(f"ERRO ao processar trecho {t}: {str(e)}")
    
    print("-" * 90)
    print(f"CUSTO TOTAL DO SISTEMA: R$ {custo_total:.2f}")
    print("=" * 90)

def _ler_diametros_colados():
    """
    Lê uma coluna colada do Excel (um diâmetro por linha) ou qualquer mistura de separadores.
    Termine a colagem com uma linha vazia (Enter duas vezes).
    Retorna lista de floats.
    """
    print("\nCole a coluna de diâmetros (um por linha) e pressione Enter duas vezes para finalizar:")
    linhas_lidas = []
    while True:
        try:
            l = input()
        except EOFError:
            break
        if l.strip() == "":
            break
        linhas_lidas.append(l)

    bruto = "\n".join(linhas_lidas)

    # Extrai números na ordem, aceitando 10, 10.0, 10,5 etc.
    import re
    tokens = re.findall(r"[-+]?\d+(?:[.,]\d+)?", bruto)
    diametros = [float(t.replace(",", ".")) for t in tokens]
    return diametros

def realizar_orcamento_manual(linhas, registros_sinapi, tes_encontradas, solucao_otima=None, velocidade_maxima_do_teste=2.0, registros_vazoes=None):
    print("\nGostaria de realizar o orçamento para diâmetros diferentes daqueles escolhidos pela otimização? (responda com sim ou não)")
    resposta = input().strip().lower()
    if resposta != "sim":
        return

    # NOVO: ler colagem da coluna do Excel (um por linha), finalizar com linha vazia
    diametros_escolhidos = _ler_diametros_colados()
    if not diametros_escolhidos:
        print("Nenhum diâmetro informado. Operação cancelada.")
        return

    # Cálculo “sem reduções” e detalhamento (suas funções já adicionadas)
    custo_total_manual, detalhamento = calcular_orcamento_manual(
        diametros_escolhidos, linhas, registros_sinapi, tes_encontradas
    )
    print("\n=== Orçamento Manual (sem reduções) ===")
    print(f"Custo total: R$ {custo_total_manual:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    print("Itens (ID | Z | Diâmetro | Custo | Origem):")
    for it in detalhamento:
        print(f"{it['id']} | {it['z']:.2f} | {it['diam']:.1f} | R$ {it['custo']:,.2f} | {it['origem']}".replace(",", "X").replace(".", ",").replace("X", "."))

    if solucao_otima is not None:
        dados_otimo = coletar_dados_por_diametro(solucao_otima, linhas)  # com reduções
        dados_manual = coletar_dados_orcamento_manual(diametros_escolhidos, linhas, registros_sinapi, tes_encontradas)  # sem reduções
        criar_graficos_comparativos(dados_otimo, dados_manual)

        velocidades_otimo = coletar_velocidades_solucao(solucao_otima, linhas)
        velocidades_manual = coletar_velocidades_manual(diametros_escolhidos, linhas, registros_vazoes)
        criar_histograma_velocidades(velocidades_otimo, velocidades_manual, velocidade_maxima_do_teste)

def coletar_dados_por_diametro(solucao, linhas):
    """
    Agrega comprimento e custo por diâmetro da solução ótima (com reduções já refletidas em 'preço total').
    """
    from collections import defaultdict
    dados = defaultdict(lambda: {'comprimento': 0.0, 'custo': 0.0})
    line_map = {(ln.get("id", ln.get("index_interno")), ln.get("inicio")): ln for ln in linhas}

    for key, j in solucao.items():
        if key not in line_map:
            continue
        ln = line_map[key]
        try:
            diam = float(ln["diâmetros nominais adotados:"][j-1])
            custo = float(ln["preço total"][j-1])           # com reduções (estado corrente)
            L = float(ln.get("comprimento", 0.0))
            dados[diam]['comprimento'] += L
            dados[diam]['custo'] += custo
        except Exception:
            continue
    return dict(dados)

def calcular_orcamento_manual(diametros_escolhidos, linhas, registros_sinapi, tes_encontradas):
    """
    Calcula custo total e detalhamento por linha do orçamento manual sem reduções.
    Não muta as linhas.
    """
    # Ordenação (ID asc, Z desc)
    ordered_lines = []
    for ln in linhas:
        ident = ln.get("id") if ln.get("id") is not None else ln.get("index_interno")
        z_coord = ln.get("inicio")[2] if ln.get("inicio") and len(ln.get("inicio")) > 2 else -float('inf')
        ordered_lines.append((ident, z_coord, ln))
    ordered_lines.sort(key=lambda x: (x[0] is None, x[0], -x[1]))

    custo_total = 0.0
    detalhamento = []
    n = min(len(diametros_escolhidos), len(ordered_lines))
    for i in range(n):
        ident, zc, ln = ordered_lines[i]
        diam = float(diametros_escolhidos[i])

        _garantir_baseline_orcamento(ln)
        diams_ln = [float(d) for d in ln.get("diâmetros nominais adotados:", [])]
        pos = next((j for j, d in enumerate(diams_ln) if abs(d - diam) < 1e-6), -1)

        if pos >= 0 and pos < len(ln["preço_total_original"]):
            preco = float(ln["preço_total_original"][pos])
            origem = "baseline"
        else:
            preco = float(_preco_manual_por_componentes(ln, diam, registros_sinapi, tes_encontradas))
            origem = "componentes"

        custo_total += preco
        detalhamento.append({"id": ident, "z": zc, "diam": diam, "custo": preco, "origem": origem})

    return custo_total, detalhamento

def _garantir_baseline_orcamento(linha):
    """
    Garante baseline para o orçamento manual sem colidir com o helper do pipeline.
    Se _garantir_baseline existir (pipeline), usa-a; caso contrário, cria snapshots locais.
    """
    try:
        _garantir_baseline(linha)  # se existir no pipeline
    except NameError:
        if "preço_total_original" not in linha:
            linha["preço_total_original"] = list(linha.get("preço total", []))
        if "perda_carga_original" not in linha:
            linha["perda_carga_original"] = list(linha.get("perda de carga", []))


def _preco_manual_por_componentes(ln, diametro, registros_sinapi, tes_encontradas):
    """
    Calcula custo sem reduções (tubo, TE, joelhos, hidr, rgl, rg) para um diâmetro arbitrário,
    sem mutar ln e sem ler 'preço total'.
    """
    L = float(ln.get("comprimento", 0.0))
    total = 0.0
    total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "tubo", L))
    if pertence_ao_te(ln, tes_encontradas):
        total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "te", L))
    if int(ln.get("joelho_45", 0)) == 1:
        total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "joelho 45", L))
    if int(ln.get("joelho_90", 0)) == 1:
        total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "joelho 90", L))
    textos = [str(t).lower() for t in ln.get("textos_associados", [])]
    if any(t.startswith("hidr") for t in textos):
        total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "hidr", L))
    if "rgl" in textos:
        total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "rgl", L))
    if "rg" in textos:
        total += float(buscar_preco_por_diametro(registros_sinapi, diametro, "rg", L))
    return total

def coletar_dados_orcamento_manual(diametros_escolhidos, linhas, registros_sinapi, tes_encontradas):
    """
    Agrega comprimento e custo por diâmetro para o orçamento manual sem reduções.
    Usa baseline se o diâmetro existir na linha; caso contrário calcula por componentes.
    """
    from collections import defaultdict
    dados_por_diametro = defaultdict(lambda: {'comprimento': 0.0, 'custo': 0.0})

    # Mesma ordenação usada na UI do orçamento (ID asc, Z desc)
    ordered_lines = []
    for ln in linhas:
        ident = ln.get("id") if ln.get("id") is not None else ln.get("index_interno")
        z_coord = ln.get("inicio")[2] if ln.get("inicio") and len(ln.get("inicio")) > 2 else -float('inf')
        ordered_lines.append((ident, z_coord, ln))
    ordered_lines.sort(key=lambda x: (x[0] is None, x[0], -x[1]))

    n = min(len(diametros_escolhidos), len(ordered_lines))
    for i in range(n):
        ident, z_coord, ln = ordered_lines[i]
        diametro = float(diametros_escolhidos[i])
        L = float(ln.get("comprimento", 0.0))

        _garantir_baseline_orcamento(ln)
        diams_ln = [float(d) for d in ln.get("diâmetros nominais adotados:", [])]
        pos = next((j for j, d in enumerate(diams_ln) if abs(d - diametro) < 1e-6), -1)

        if pos >= 0 and pos < len(ln["preço_total_original"]):
            custo = float(ln["preço_total_original"][pos])   # baseline sem reduções
        else:
            custo = float(_preco_manual_por_componentes(ln, diametro, registros_sinapi, tes_encontradas))

        dados_por_diametro[diametro]['comprimento'] += L
        dados_por_diametro[diametro]['custo'] += custo

    return dict(dados_por_diametro)

def criar_graficos_comparativos(dados_otimo, dados_manual):
    """
    Cria gráficos comparativos entre solução ótima e orçamento manual.
    
    Args:
        dados_otimo: dados por diâmetro da solução ótima
        dados_manual: dados por diâmetro do orçamento manual
    """
    # Obter todos os diâmetros únicos
    todos_diametros = set(dados_otimo.keys()) | set(dados_manual.keys())
    diametros_ordenados = sorted(todos_diametros)
    
    # Preparar dados para os gráficos
    comprimentos_otimo = [dados_otimo.get(d, {'comprimento': 0.0})['comprimento'] for d in diametros_ordenados]
    comprimentos_manual = [dados_manual.get(d, {'comprimento': 0.0})['comprimento'] for d in diametros_ordenados]
    
    custos_otimo = [dados_otimo.get(d, {'custo': 0.0})['custo'] for d in diametros_ordenados]
    custos_manual = [dados_manual.get(d, {'custo': 0.0})['custo'] for d in diametros_ordenados]
    
    # Configurar a figura com 2 subplots
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(15, 6))
    
    # Configurar posições das barras
    x = np.arange(len(diametros_ordenados))
    width = 0.35
    
    # Gráfico 1: Comprimentos por Diâmetro
    bars1_otimo = ax1.bar(x - width/2, comprimentos_otimo, width, label='Solução Ótima', 
                         color='skyblue', alpha=0.8)
    bars1_manual = ax1.bar(x + width/2, comprimentos_manual, width, label='Orçamento Manual', 
                          color='lightcoral', alpha=0.8)
    
    ax1.set_xlabel('Diâmetro (mm)')
    ax1.set_ylabel('Comprimento Total (m)')
    ax1.set_title('Comprimento Total de Tubos por Diâmetro')
    ax1.set_xticks(x)
    ax1.set_xticklabels([f'{int(d)}' for d in diametros_ordenados])
    ax1.legend()
    ax1.grid(True, alpha=0.3)
    
    # Adicionar valores nas barras do gráfico 1
    for bar in bars1_otimo:
        height = bar.get_height()
        if height > 0:
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{height:.1f}m', ha='center', va='bottom', fontsize=8)
    
    for bar in bars1_manual:
        height = bar.get_height()
        if height > 0:
            ax1.text(bar.get_x() + bar.get_width()/2., height + 0.1,
                    f'{height:.1f}m', ha='center', va='bottom', fontsize=8)
    
    # Gráfico 2: Custos por Diâmetro
    bars2_otimo = ax2.bar(x - width/2, custos_otimo, width, label='Solução Ótima', 
                         color='lightgreen', alpha=0.8)
    bars2_manual = ax2.bar(x + width/2, custos_manual, width, label='Orçamento Manual', 
                          color='orange', alpha=0.8)
    
    ax2.set_xlabel('Diâmetro (mm)')
    ax2.set_ylabel('Custo Total (R$)')
    ax2.set_title('Custo Total por Diâmetro')
    ax2.set_xticks(x)
    ax2.set_xticklabels([f'{int(d)}' for d in diametros_ordenados])
    ax2.legend()
    ax2.grid(True, alpha=0.3)
    
    # Adicionar valores nas barras do gráfico 2
    for bar in bars2_otimo:
        height = bar.get_height()
        if height > 0:
            ax2.text(bar.get_x() + bar.get_width()/2., height + 20,
                    f'R${height:.0f}', ha='center', va='bottom', fontsize=8)
    
    for bar in bars2_manual:
        height = bar.get_height()
        if height > 0:
            ax2.text(bar.get_x() + bar.get_width()/2., height + 20,
                    f'R${height:.0f}', ha='center', va='bottom', fontsize=8)
    
    # Adicionar resumo dos totais
    total_custo_otimo = sum(custos_otimo)
    total_custo_manual = sum(custos_manual)
    total_comprimento_otimo = sum(comprimentos_otimo)
    total_comprimento_manual = sum(comprimentos_manual)
    
    fig.suptitle(f'Comparação: Solução Ótima vs Orçamento Manual\n'
                f'Custo Total - Ótimo: R${total_custo_otimo:.2f} | Manual: R${total_custo_manual:.2f} | '
                f'Diferença: R${total_custo_manual - total_custo_otimo:.2f}\n'
                f'Comprimento Total - Ótimo: {total_comprimento_otimo:.1f}m | Manual: {total_comprimento_manual:.1f}m',
                fontsize=12)
    
    plt.tight_layout()
    plt.subplots_adjust(top=0.85)
    plt.show()

def coletar_velocidades_solucao(solucao, linhas):
    """
    Coleta as velocidades finais para cada trecho de uma solução ótima.

    Args:
        solucao (dict): Dicionário da solução ótima { (id, inicio): j }.
        linhas (list): A lista completa de dados das linhas.

    Returns:
        list: Uma lista com os valores de velocidade da solução.
    """
    velocidades = []
    line_map = { (ln.get("id", ln.get("index_interno")), ln.get("inicio")): ln for ln in linhas }

    for key, j_idx in solucao.items():
        if key in line_map:
            linha = line_map[key]
            try:
                # O índice 'j' do Pyomo é 1-based, então acessamos a lista com j_idx-1
                velocidade_escolhida = float(linha["velocidade fluido (m/s)"][j_idx - 1])
                velocidades.append(velocidade_escolhida)
            except (IndexError, ValueError, TypeError):
                # Ignora se houver algum erro de dado para esta linha
                continue
    return velocidades

def coletar_velocidades_manual(diametros_escolhidos, linhas, registros_vazoes):
    """
    Coleta as velocidades para a solução manual.
    - Se o diâmetro manual existir entre os candidatos do trecho, usa a velocidade pré-calculada.
    - Caso contrário, calcula v = Q/A usando a vazão do trecho e a área (m^2) obtida em registros_vazoes.
    """
    velocidades = []

    # Ordenar linhas exatamente como no orçamento manual (ID asc, Z desc)
    ordered_lines = []
    for ln in linhas:
        ident = ln.get("id") if ln.get("id") is not None else ln.get("index_interno")
        z_coord = ln.get("inicio")[2] if ln.get("inicio") and len(ln.get("inicio")) > 2 else -float('inf')
        ordered_lines.append((ident, z_coord, ln))
    ordered_lines.sort(key=lambda x: (x[0] is None, x[0], -x[1]))

    # Helper local para buscar área (m^2) por diâmetro nominal (mm) em registros_vazoes
    def _buscar_area_por_diametro_mm(d_mm, registros, tol=1e-6):
        if registros is None:
            return None
        try:
            dmm = float(d_mm)
        except:
            return None
        for reg in registros:
            try:
                dn = float(reg.get("diam_nom"))
                if abs(dn - dmm) < tol:
                    a = reg.get("area")
                    return float(a) if a is not None else None
            except:
                continue
        return None

    for i, (_, _, ln) in enumerate(ordered_lines):
        if i >= len(diametros_escolhidos):
            break

        diametro_manual = diametros_escolhidos[i]
        try:
            # 1) Tentar achar nos candidatos do trecho
            idx = -1
            for k, d_cand in enumerate(ln.get("diâmetros nominais adotados:", [])):
                try:
                    if abs(float(d_cand) - float(diametro_manual)) < 1e-6:
                        idx = k
                        break
                except:
                    continue

            if idx != -1:
                # velocidade já pré-computada na mesma posição
                vel_list = ln.get("velocidade fluido (m/s)", [])
                if 0 <= idx < len(vel_list):
                    velocidades.append(float(vel_list[idx]))
                else:
                    # fallback raro: posição não existe
                    velocidades.append(0.0)
                continue

            # 2) Caso o diâmetro não exista nos candidatos do trecho, calcular v = Q/A
            Q = float(ln.get("vazao_m3_s", 0.0))  # vazão do próprio trecho
            A = _buscar_area_por_diametro_mm(diametro_manual, registros_vazoes)
            if A is not None and A > 0:
                velocidades.append(Q / A)
            else:
                velocidades.append(0.0)

        except (IndexError, ValueError, TypeError):
            velocidades.append(0.0)

    return velocidades

def criar_histograma_velocidades(velocidades_otimo, velocidades_manual, v_max):
    """
    Cria e exibe um histograma comparativo da distribuição de velocidades.

    Args:
        velocidades_otimo (list): Lista de velocidades da solução ótima.
        velocidades_manual (list): Lista de velocidades da solução manual.
        v_max (float): Velocidade máxima usada como parâmetro na otimização.
    """
    plt.figure(figsize=(12, 7))
    
    # Define os "bins" (faixas) do histograma. De 0 a v_max + 0.5 m/s, com passos de 0.1 m/s
    bins = np.arange(0, v_max + 0.6, 0.1)

    plt.hist(velocidades_otimo, bins=bins, alpha=0.7, label=f'Solução Ótima (V_máx = {v_max} m/s)', color='skyblue', edgecolor='black', rwidth=0.85)
    plt.hist(velocidades_manual, bins=bins, alpha=0.7, label='Orçamento Manual', color='salmon', edgecolor='black', rwidth=0.85)

    # Adiciona uma linha vertical para a velocidade máxima
    plt.axvline(x=v_max, color='red', linestyle='--', linewidth=2, label=f'Limite V_máx = {v_max} m/s')

    plt.xlabel('Velocidade do Fluido (m/s)', fontsize=12)
    plt.ylabel('Número de Trechos', fontsize=12)
    plt.title('Distribuição das Velocidades nos Trechos da Rede', fontsize=14, weight='bold')
    plt.legend()
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.tight_layout()
    plt.show()


def gerar_histograma_margens_seguranca(margens_de_seguranca):
    """
    Gera histograma das margens de segurança.
    
    Args:
        margens_de_seguranca: Lista com os valores das margens de segurança já calculadas.
    """
    if not margens_de_seguranca:
        print("\nAVISO: Nenhuma margem de segurança fornecida para o histograma.")
        return
    
    print("\n" + "=" * 90)
    print("HISTOGRAMA DAS MARGENS DE SEGURANÇA".center(90))
    print("=" * 90)
    
    # Estatísticas
    print(f"\nEstatísticas das Margens de Segurança:")
    print(f"  Total de caminhos: {len(margens_de_seguranca)}")
    print(f"  Margem mínima: {min(margens_de_seguranca):.4f} m.c.a.")
    print(f"  Margem máxima: {max(margens_de_seguranca):.4f} m.c.a.")
    print(f"  Margem média: {np.mean(margens_de_seguranca):.4f} m.c.a.")
    print(f"  Desvio padrão: {np.std(margens_de_seguranca):.4f} m.c.a.")
    
    # --- Criar Histograma ---
    plt.figure(figsize=(12, 7))
    
    # Definir bins dinamicamente
    min_val = np.floor(min(margens_de_seguranca) * 2) / 2
    max_val = np.ceil(max(margens_de_seguranca) * 2) / 2
    bins = np.arange(min_val, max_val + 0.5, step=0.5)
    
    # Plotar histograma
    plt.hist(margens_de_seguranca, 
             bins=bins, 
             color='skyblue',
             edgecolor='black',
             alpha=0.75,
             rwidth=0.85)
    
    # Linha de referência em margem = 0
    plt.axvline(0, color='red', linestyle='--', linewidth=2, 
                label='Margem de Segurança = 0 m.c.a.')
    
    # Configurações do gráfico
    plt.title('Distribuição das Margens de Segurança', fontsize=15, weight='bold')
    plt.xlabel('Margem de Segurança (m.c.a.)', fontsize=13)
    plt.ylabel('Número de Caminhos', fontsize=13)
    plt.legend(fontsize=11)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.grid(axis='x', linestyle=':', alpha=0.5)
    plt.tight_layout()
    
    # Salvar e exibir
    plt.savefig("histograma_margens_seguranca.png", dpi=300, bbox_inches='tight')
    print(f"\n✓ Histograma salvo como 'histograma_margens_seguranca.png'")
    plt.show()
    
    print("=" * 90)


def main():
    caminho_arquivo_dxf = ""
    caminho_peso_relativo = ""
    caminho_vazoes = ""
    caminho_perda = ""
    registros_sinapi = pd.read_excel("").to_dict(orient="records")
    registros_reducao = carregar_planilha_reducao("")


    # 1) Ler planilha de pesos relativos
    tabela_pesos_relativos = carregar_planilha_pesos_relativos(caminho_peso_relativo)
    print("== Planilha de Pesos Relativos (dados carregados) ==")
    for idx, lp in enumerate(tabela_pesos_relativos, start=1):
        print(f"{idx}) Aparelho: {lp.get('aparelho_sanitario', '')}, "
              f"Peça: {lp.get('peca_de_utilizacao', '')}, "
              f"Sigla: {lp.get('sigla', '')}, "
              f"Vazão (m³/s): {lp.get('vazao_de_projeto_m3_s', '')}, "
              f"Peso Relativo: {lp.get('peso_relativo', 0)}")

    print("== Planilha de perda de carga redução (dados carregados) ==")
    for idx, lp in enumerate(registros_reducao, start=1):
        print(f"{idx}) diametro entrada: {lp.get('dia_nom_entrada', '')}, "
              f"diametro saida: {lp.get('dia_nom_saida', '')}, "
              f"coeficiente: {lp.get('coeficiente', '')}, ")

    # 2) Ler DWG/DXF
    try:
        doc = ezdxf.readfile(caminho_arquivo_dxf)
    except Exception as e:
        print(f"Não foi possível abrir o arquivo: {e}")
        return
    msp = doc.modelspace()

    # 3) Ler textos e linhas
    textos_msp = ler_textos_e_mtexts(msp)
    linhas_msp = ler_linhas(msp)

    # 4) Associar textos às linhas
    associar_identificadores(linhas_msp, textos_msp)
    associar_textos_nao_numericos(linhas_msp, textos_msp)

    # 5) Associa pesos relativos
    associar_peso_relativo(linhas_msp, tabela_pesos_relativos)

    # 6) Identificar TEs
    tes_encontradas = identificar_tes(linhas_msp)

    # 7) Construir caminhos das siglas até 'res'
    caminhos_siglas = construir_caminhos_siglas_para_res(linhas_msp)

    # 8) Somar pesos relativos nos percursos
    somar_pesos_relativos(linhas_msp, caminhos_siglas)

    # 9) Processar TE Passagem Direta ou Saída Lateral
    processar_te_passagem_lateral(linhas_msp, tes_encontradas, caminhos_siglas)

    # 10) Processar Joelhos (45 e 90)
    processar_joelhos(linhas_msp, caminhos_siglas)

    # 11) Calcular pressão estática e perda de carga máxima admissível para cada sigla
    calcular_pressao_estatica_e_perda_carga(linhas_msp, caminhos_siglas, textos_msp, tabela_pesos_relativos)

    # 12) Calcular vazão m^3/s para cada linha
    calcular_vazao(linhas_msp)

    # 13) Carregar a planilha de vazões máximas e calcular diâmetros adotados
    registros_vazoes = carregar_planilha_vazoes_maximas(caminho_vazoes)
    calcular_diametros_adotados(linhas_msp, registros_vazoes)

    # 14) Carregar a planilha de "perda de carga localizada" e calcular comprimentos equivalentes
    registros_perda = carregar_planilha_perda_de_carga(caminho_perda)
    calcular_comprimentos_equivalentes(linhas_msp, registros_perda)

    # 15) Calcular velocidade fluido (m/s) para cada linha
    calcular_velocidade_fluido(linhas_msp)

    # 16) Calcular Reynolds para cada linha
    calcular_reynolds(linhas_msp)

    # 17) Calcular fator de atrito para cada linha
    calcular_fator_atrito(linhas_msp)

    # 18) Calcular perda de carga unitária para cada linha
    calcular_perda_carga_unitaria(linhas_msp)

    # 19) Calcular comprimento virtual (m) para cada linha
    calcular_comprimento_virtual(linhas_msp)

    # 20) Calcular perda de carga (m) para cada linha
    calcular_perda_carga(linhas_msp)
    
    # 21) Calcular perda de carga hidrômetro para as linhas (apenas para aquelas com "hidr..." nos textos associados)
    calcular_perda_carga_hidrometro(linhas_msp)
    atualizar_perda_carga_com_hidrometro(linhas_msp)

    # 22) Calcular preço por diâmetro e preço total para cada linha usando dados sinapi
    calcular_preco_diametro(linhas_msp, registros_sinapi, tes_encontradas)

    # 23) Carregar a planilha de perda de carga redução e calcular perda de carga redução para cada sigla
    calcular_perda_carga_reducao(caminhos_siglas, registros_reducao, linhas_msp)

    # 24) Após calcular a perda de carga redução, atualize os preços
    atualizar_preco_perda_reducao(registros_sinapi, linhas_msp)

    # 25) Remove as possibilidades de redução que são inválidas
    filtrar_perda_carga_reducao(linhas_msp)
    
    # organizando os dados antes de chamar solve_linear_model
    linhas_map = {}
    for ln in linhas_msp:
        ident = ln.get("id", ln.get("index_interno"))
        inicio = ln.get("inicio")
        key = (ident, inicio)
        linhas_map[key] = ln

    
    model, sol = solve_linear_model(linhas_msp, caminhos_siglas)

    
    if sol:  # Se foi encontrada uma solução viável
        # Ajustar diâmetros para eliminar margens negativas
        print("\nVerificando margens de segurança e ajustando diâmetros se necessário...")
        
        
        # Atualizar custos e perdas de carga finais
        update_costs_and_losses_for_solution(sol, linhas_msp, caminhos_siglas, linhas_map)

        # Ajustar diâmetros começando SEMPRE da última redução antes da sigla
        sol = ajustar_diametros_para_margens_negativas(sol, linhas_msp, caminhos_siglas)

        
        # Imprimir resultados
        margens = imprimir_resultados_por_sigla(caminhos_siglas, sol, linhas_msp)
        imprimir_resultados_otimizacao(sol, linhas_msp)

        gerar_histograma_margens_seguranca(margens)
        
        # Adicionar chamada para orçamento manual com solução ótima para comparação
        velocidade_maxima_do_teste = 3.0
        realizar_orcamento_manual(linhas_msp, registros_sinapi, tes_encontradas, sol, velocidade_maxima_do_teste=velocidade_maxima_do_teste)
    else:
        print("Não foi possível encontrar uma solução viável. Verifique os dados de entrada.")


    # ----------------- Imprimir Resultados Finais ----------------- #
    print("\n== LINHAS ENCONTRADAS NO MODEL SPACE ==")

    # Criar lista ordenável
    ordered_lines = []
    for ln in linhas_msp:
        ident = ln["id"] if ln["id"] is not None else ln["index_interno"]
        # Extrair coordenada Z
        z_coord = ln["inicio"][2] if ln["inicio"] and len(ln["inicio"]) > 2 else -float('inf')
        ordered_lines.append((ident, z_coord, ln))

    # Ordenar por ID e depois por Z decrescente
    ordered_lines.sort(key=lambda x: (x[0] is None, x[0], -x[1]))  # None vai para o final, Z em ordem decrescente

    # Iterar pela lista ordenada
    for i, (ident, z_coord, ln) in enumerate(ordered_lines, start=1):
        print(f"{i}) Linha: {ident}")
        print(f"   Início: {ln['inicio']}, Fim: {ln['fim']}")
        print(f"   Comprimento: {ln['comprimento']:.2f}, Ângulo: {ln['angulo']:.2f}°")
        print(f"   Joelho 45°: {ln['joelho_45']}")
        print(f"   Joelho 90°: {ln['joelho_90']}")
        print(f"   TE Saída Lateral: {ln['te_saida_lat']}")
        print(f"   TE Passagem Direta: {ln['te_pass_dir']}")
        if ln["id"] is not None:
            print(f"   Identificador numérico no desenho: {ln['id']}")
        if ln["textos_associados"]:
            print(f"   Textos associados: {ln['textos_associados']}")
        print(f"   Peso relativo total: {ln['peso_relativo_total']:.2f}")
        print(f"   Vazão m^3/s: {ln['vazao_m3_s']:.10f}")
        print(f"   Diâmetros nominais adotados: {ln.get('diâmetros nominais adotados:', [])}")
        print(f"   Diâmetro interno (m): {ln.get('diâmetro interno (m):', [])}")
        print(f"   Área (m^2): {ln.get('área (m^2):', [])}")
        print(f"   Comprimentos equivalentes: {ln.get('comprimentos equivalentes', [])}")
        print(f"   Comprimento virtual (m): {ln.get('comprimento virtual (m)', [])}")
        print(f"   Velocidade fluido (m/s): {ln.get('velocidade fluido (m/s)', [])}")
        print(f"   Reynolds: {ln.get('Reynolds', [])}")
        print(f"   Fator de atrito: {ln.get('fator de atrito', [])}")
        print(f"   Perda de carga unitária: {ln.get('perda de carga unitária', [])}")
        print(f"   Perda de carga: {ln.get('perda de carga', [])}")
        print(f"   Perda de carga hidrômetro: {ln.get('perda de carga hidrômetro', [])}")
        print(f"   Preço por diâmetro: {ln.get('preço por diâmetro', [])}")
        print(f"   Preço total: {ln.get('preço total', [])}")
        print(f"   Perda de carga redução: {ln.get('perda de carga redução', [])}")

    print("\n== CONEXÕES DO TIPO TE ENCONTRADAS ==")
    if not tes_encontradas:
        print("Nenhuma conexão TE encontrada.")
    else:
        for idx, te in enumerate(tes_encontradas, start=1):
            print(f"{idx}) Coordenada TE: {te['coord']}, Linhas: {te['line_ids']}")

    print("\n== CAMINHOS DAS SIGLAS ATÉ O RESERVATÓRIO ==")
    if not caminhos_siglas:
        print("Nenhum caminho encontrado.")
    else:
        for sigla, info in caminhos_siglas.items():
            print(f"Sigla: {info['sigla_completa']}")
            print(f"  Linhas TE relacionadas: {info['linhas_tes']}")
            print(f"  Pressão estática: {info['pressao_estatica']:.2f}")
            print(f"  Perda de carga máxima admissível: {info['perda_carga_max_adm']:.2f}")
            if info.get('msg', ""):
                print(f"  ALERTA: {info['msg']}")
            print("  Segmentos no caminho (line_id, angle, peso_relativo, coordenadas_iniciais):")
            for seg in info['segmentos_no_caminho']:
                print(f"    - ID: {seg['line_id']}, Ângulo: {seg['angle']:.2f}°, Peso: {seg['peso_relativo']}, Coord Ini: {seg['coordenadas_iniciais']}")
            print("----")

if __name__ == "__main__":
    main()

