import re
from string import ascii_uppercase as Alfabeto
from typing import Literal

import gspread
from gspread.worksheet import Worksheet

gc = gspread.service_account()
sh = gc.open_by_key("1tv2j_bPKGiKKZI7iAfYxV7YyiHpEq4nD3MQsatKgNOw") 

worksheet = sh.get_worksheet(0)


formula_obj = "=SEERRO(SOMA('{titulo}'!{celulas})/SOMA({total}))"
formula_dis = "=SEERRO((SOMA('{titulo}'!{celulas}) + SOMA(DESLOC('{titulo}'!{celulas};1;)) * 0,5) /SOMA({total}))"
formula_rede = "=SE(SOMA({linha_atual})<>0;{formula};"")"

def escreve_percurso(ponteiro: str, celula_resultados: str, escolas: list, mapeamento: Worksheet, resultado: Worksheet ):
    titulo = resultado.title
    
    for i, turmas in enumerate(escolas):
        if turmas > 1:
            celulas = celula_resultados + ':' + mover_celula(celula_resultados, 'coluna', turmas-1)
        else:
            celulas = celula_resultados
        
        total = f"'{titulo}'!" + mover_celula(celulas,'linha', 15)

        # Escreve fórmulas das questões objetivas
        formulas = []
        ponteiro_inicial = ponteiro
        for _ in range(5):
            #print(ponteiro)
            formulas.append([formula_obj.format(titulo=titulo, celulas=celulas, total=total)])
            celulas = mover_celula(celulas, 'linha', 1)
            ponteiro = mover_celula(ponteiro,'linha', 1)

        # Escreve as fórmulas das questões abertas
        for _ in range(5):
            #print(ponteiro)
            formulas.append([formula_dis.format(titulo=titulo, celulas=celulas, total=total)])
            celulas = mover_celula(celulas, 'linha', 2)
            ponteiro = mover_celula(ponteiro,'linha', 1)
        intervalo = ponteiro_inicial + ':' + mover_celula(ponteiro, 'linha', -1)
        mapeamento.update(formulas, intervalo, raw=False)

        ponteiro = mover_celula(ponteiro,'linha', -10)
        ponteiro = mover_celula(ponteiro,'coluna', 1)
        celula_resultados = mover_celula(celula_resultados, 'coluna', turmas)


def escritor(ponteiro: str, celula_resultados: str, escolas: list, mapeamento: Worksheet, resultado: Worksheet ) -> None:
    '''
    Escreve as fórmulas nas abas de mapeamento das planilhas

    Args:
        ponteiro (str): Posição onde irá ser escrita a fórmula
        celula_resultados (str): Posição inicial da escola nos resultados
        escolas (list): Lista com o número de turmas em cada escola
        mapeamento (Worksheet): Planilha de mapeamento 
        resultado (Worksheet): Planilha de resultado

    Return:
        None

    Essa função não retorna nada, ela apenas afeta as planilhas de mapeamento passadas como argumento
    '''
    titulo = resultado.title
    #escolas = [4,1,2,1,2]

    for turmas in escolas:
        if turmas > 1:
            celulas = celula_resultados + ':' + mover_celula(celula_resultados, 'coluna', turmas-1)
        else:
            celulas = celula_resultados
        total = f"'{titulo}'!" + mover_celula(celulas,'linha', 17)
        
        # Escreve fórmulas das questões objetivas
        formulas = []
        ponteiro_inicial = ponteiro
        for _ in range(7):
            formulas.append([formula_obj.format(titulo=titulo, celulas=celulas, total=total)])
            celulas = mover_celula(celulas, 'linha', 1)
            ponteiro = mover_celula(ponteiro,'linha', 1)

        # Escreve as fórmulas das questões abertas    
        for _ in range(5):
            formulas.append([formula_dis.format(titulo=titulo, celulas=celulas, total=total)])
            celulas = mover_celula(celulas, 'linha', 2)
            ponteiro = mover_celula(ponteiro,'linha', 1)
        intervalo = ponteiro_inicial + ':' + mover_celula(ponteiro, 'linha', -1)
        mapeamento.update(formulas, intervalo, raw=False)

        ponteiro = mover_celula(ponteiro,'linha', -12)
        ponteiro = mover_celula(ponteiro,'coluna', 1)
        celula_resultados = mover_celula(celula_resultados, 'coluna', turmas)

    # Escrever as fórmulas da rede
    celulas = mover_celula(celula_resultados, 'coluna', -sum(escolas)) + ':' + mover_celula(celula_resultados, 'coluna', -1)
    linha_atual =  mover_celula(ponteiro, 'coluna', -5) + ':' + mover_celula(ponteiro, 'coluna', -1)
    total = f"'{titulo}'!" + mover_celula(celulas,'linha', 17)

    formulas = []
    ponteiro_inicial = ponteiro
    for _ in range(7):
        formulas.append([formula_rede.format(linha_atual=linha_atual, formula=formula_obj.format(titulo=titulo, celulas=celulas, total=total)[8:-1])])
        celulas = mover_celula(celulas, 'linha', 1)
        ponteiro = mover_celula(ponteiro,'linha', 1)
        linha_atual = mover_celula(linha_atual, 'linha', 1)
    for _ in range(5):
        formulas.append([formula_rede.format(linha_atual=linha_atual, formula=formula_dis.format(titulo=titulo, celulas=celulas, total=total)[8:-1])])
        celulas = mover_celula(celulas, 'linha', 2)
        ponteiro = mover_celula(ponteiro,'linha', 1)
        linha_atual = mover_celula(linha_atual, 'linha', 1)
    intervalo = ponteiro_inicial + ':' + mover_celula(ponteiro, 'linha', -1)
    mapeamento.update(formulas, intervalo, raw=False)

def mover_celula(endereco: str, direcao: Literal['linha', 'coluna'], deslocamento: int) -> str:
    '''
    Move os endereços de células das planilhas

    Args:
        endereco (str): Endereço inicial que desejo mover
        direcao (str): A direção do movimento que deve ser 'linha' ou 'coluna'
        deslocamento (int): O quanto a célula irá se deslocar, este valor pode ser positivo ou negativo
    
    Return:
        str: A função retorna o novo endereço depois de realizar o movimento
    '''
    if ':' in endereco:
        i, j = endereco.split(':')
        return mover_celula(i, direcao, deslocamento) + ':' + mover_celula(j, direcao, deslocamento)     

    padrao = r"([A-Z]+)(\d+)"
    resultado = re.match(padrao, endereco)
    coluna, linha = resultado.groups()
    linha = int(linha)

    if direcao == 'linha':
        linha += deslocamento
    elif direcao == 'coluna':
        coluna =  move_coluna(coluna, deslocamento)
    else:
        raise ValueError('Comando inválido, o valor da variável direcao deve ser "linha" ou "coluna"')
    return coluna + str(linha)

def move_coluna(coluna, des):
    if len(coluna) > 1:
        indice_coluna =  converte_letra(coluna[0]) * 26 + converte_letra(coluna[1])
    else:
        indice_coluna =  converte_letra(coluna)
    
    indice_coluna += des

    if indice_coluna > 26:
        i = indice_coluna // 26
        j = indice_coluna % 26
        return Alfabeto[i - 1] + Alfabeto[j - 1]
    return Alfabeto[indice_coluna - 1]

def converte_letra(letra):
    return Alfabeto.index(letra) + 1

if __name__ == '__main__':
    escritor('I21', 'D21' ,[4,1,2,1,2],  sh.get_worksheet(0), sh.get_worksheet(1) )