from time import sleep
from escritor import escritor, escreve_percurso, mover_celula
import gspread

gc = gspread.service_account()
sh = gc.open_by_key("1wI8OmJRFtVwcYHuuBxS28dzLH1IAoVNJK546UQYMtAg") 

mapeamento = sh.get_worksheet(4)
resultado = sh.get_worksheet(5)

tabelas_mapeamento = mapeamento.findall('Academia')
tabelas_resultado = resultado.findall('Academia')

tabelas_mapeamento = tabelas_mapeamento[-6:]
tabelas_resultado = tabelas_resultado[-6:]

#turmas = [[4,1,2,2,3], [], [4,1,2,1,2], [], [4,1,1,1,2], [], [3,1,2,1,2], [], [4,1,1,1,3], [], [4,1,1,1,3], [], [2,1,1,1,2]]
turmas = [[4,1,1,1,3], [], [4,1,1,1,3], [], [2,1,1,1,2], []]
#turmas = [[4,1,1,1,3], [4,1,1,1,3], [2,1,1,1,2]]

print('\n',sh.title)
print('Escrevendo em: ', mapeamento.title)
for i in range(0,len(tabelas_mapeamento), 2):
#for i in range(0,len(tabelas_mapeamento)):
    ponteiro = mover_celula(tabelas_mapeamento[i].address, 'linha', 1)
    celula_resultado = mover_celula(tabelas_resultado[i].address, 'linha', 2)
    print('Celula resultado: ', celula_resultado)
    escritor(ponteiro, celula_resultado, turmas[i], mapeamento, resultado)

    ponteiro = mover_celula(tabelas_mapeamento[i+1].address, 'linha', 1)
    #ponteiro = mover_celula(ponteiro, 'coluna', 4)
    celula_resultado = mover_celula(tabelas_resultado[i+1].address, 'linha', 2)
    print('Celula resultado: ', celula_resultado)
    escreve_percurso(ponteiro, celula_resultado, turmas[i], mapeamento, resultado)
    if i==6:
        for _ in range(61, 0, -1):
            print(f'\rAguarde: {_} segundos ', end='')
            sleep(1)
        print('\r                      ')