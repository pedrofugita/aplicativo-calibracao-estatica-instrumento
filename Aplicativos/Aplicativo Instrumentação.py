import math
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
import openpyxl
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Inicializa as listas para armazenar os valores fornecidos
qi = []  # Valor verdadeiro
qo = []  # Valor indicado

# Inicializa o contador de itens adicionados
contador_itens = 0


def importar_excel():
    global contador_itens

    # Abra a caixa de diálogo para selecionar o arquivo Excel
    filepath = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])

    if filepath:
        try:
            # Carregue o arquivo Excel
            workbook = openpyxl.load_workbook(filepath)

            # Acesse a planilha desejada
            planilha = workbook.active

            # Percorra as linhas da planilha e adicione os dados à tabela
            for row in planilha.iter_rows(values_only=True):
                if len(row) >= 2:
                    valor_qi, valor_qo = map(float, row[:3])  # Converter para float com 3 casas decimais
                    tabela.insert('', 'end', values=(valor_qi, valor_qo))
                    qi.append(valor_qi)
                    qo.append(valor_qo)
                    contador_itens += 1

            # Atualize o rótulo com o número de itens adicionados
            N.config(text=f"Quantidade: {contador_itens}")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao importar dados: {str(e)}")


def adicionar_item():
    # Obtém os valores digitados nos campos de entrada
    valor_qi = qi_entry.get()
    valor_qo = qo_entry.get()

    try:
        # Converte os valores para números de ponto flutuante
        valor_qi = float(valor_qi)
        valor_qo = float(valor_qo)

        # Adiciona os valores à tabela
        tabela.insert('', 'end', values=(valor_qi, valor_qo))

        # Adiciona os valores às listas
        qi.append(valor_qi)
        qo.append(valor_qo)

        # Incrementa o contador de itens adicionados
        global contador_itens
        contador_itens += 1

        # Atualiza o rótulo com o número de itens adicionados
        N.config(text=f"Quantidade: {contador_itens}")

        # Limpa os campos de entrada após a adição
        qi_entry.delete(0, 'end')
        qo_entry.delete(0, 'end')

    except ValueError:
        # Trata exceção se a conversão para float falhar
        messagebox.showerror("Erro", "Valores inválidos. Certifique-se de usar ponto como separador decimal.")


def calculos():
    global m_valor, b_valor, sm_valor, sb_valor, sqo_valor, sqi_valor

    # Verifique se há dados suficientes para calcular m e b
    if len(qi) < 2 or len(qo) < 2:
        messagebox.showerror("Erro", "É necessário pelo menos 2 pares de valores (qi, qo) para calcular a equação da reta.")
        return

    # Cálculo de m e b
    sum_qi = sum(qi)
    sum_qo = sum(qo)
    sum_qi_qo = sum(x * y for x, y in zip(qi, qo))
    sum_qi_squared = sum(x ** 2 for x in qi)
    n = len(qi)

    m = (n * sum_qi_qo - (sum_qi) * (sum_qo)) / (n * sum_qi_squared - ((sum_qi)**2))
    b = ((sum_qo)*(sum_qi_squared)-(sum_qi_qo)*(sum_qi)) / (n * sum_qi_squared - ((sum_qi)**2))

    m_valor = m
    b_valor = b

    # Cálculo das incertezas
    soma_mb = [(m_valor * x + b_valor - y) for x, y in zip(qi, qo)]
    erros = sum(e ** 2 for e in soma_mb) / len(soma_mb)
    sqo = math.sqrt(erros)
    sm = math.sqrt((n * sqo**2) / (n * sum_qi_squared - (sum_qi)**2))
    sb = math.sqrt(((sqo**2) * sum_qi_squared) / (n * sum_qi_squared - (sum_qi)**2))
    sqi = math.sqrt((sqo**2) / (m**2))

    sqo_valor = sqo *2
    sm_valor = sm *2
    sb_valor = sb *2
    sqi_valor = sqi *2

    # Atualiza os rótulos com os valores calculados
    m_label.config(text=f"m = {m_valor:.6f} ± {sm_valor:.6f}")
    b_label.config(text=f"b = {b_valor:.6f} ± {sb_valor:.6f}")
    equation_label.config(text=f"qo = {m_valor:.3f} * qi + {b_valor:.2f}")


def grafico():
    global m_valor, b_valor, sqo_valor

    # Cria uma lista de valores preditos com base na equação da reta
    qi_valores = [x for x in range(int(min(qi)), int(max(qi)) + 1)]
    qo_predito = [m_valor * x + b_valor for x in qi_valores]

    # Cria o gráfico
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.scatter(qi, qo, label="Valor indicado pelo instrumento", color="red")

    ax.plot(qi_valores, qo_predito, label=f"qo = {m_valor:.3f}qi + {b_valor:.3f}", color="black")

    # Adicione duas linhas tracejadas
    ax.plot(qi_valores, [(m_valor * x + b_valor)+sqo_valor for x in qi_valores], linestyle='--', label="qo ± Δqo", color="blue")
    ax.plot(qi_valores, [(m_valor * x + b_valor)-sqo_valor for x in qi_valores], linestyle='--', color="blue")

    ax.set_xlabel("qi")
    ax.set_ylabel("qo")
    ax.set_title("Curva de Calibração")
    ax.legend()
    ax.grid(True)

    # Colocar gráfico na interface gráfica
    canvas = FigureCanvasTkAgg(fig, master=janela)
    canvas_widget = canvas.get_tk_widget()
    canvas_widget.grid(row=2, column=3, rowspan=7, padx=40)

    # Exibir o gráfico na interface gráfica
    canvas.draw()



def saida():
    global m_valor, b_valor, sqo_valor, bias_valor, sqi_valor
    qo = float(entrada_entry.get())
    qi = (qo - b_valor) / m_valor
    bias = qo - qi
    saida_final = qi
    bias_valor = bias



    # Atualiza os rótulos com os valores calculados
    saida_label.config(text=f"Valor verdadeiro:    {saida_final:.2f} ± {sqo_valor:.2f}")
    bias_label.config(text=f"Erros de bias:    {bias_valor:.2f}")
    imprecisao_label.config(text=f"Imprecisão:    {sqi_valor:.2f}")


# Cria a janela principal
janela = tk.Tk()
janela.title("Curva de Calibração de Instrumento")

# Cria uma tabela para a entrada de dados
tabela = ttk.Treeview(janela, columns=("Valor verdadeiro (qi)", "Valor indicado (qo)"),
                      show="headings")
tabela.heading("Valor verdadeiro (qi)", text="Valor verdadeiro (qi)")
tabela.heading("Valor indicado (qo)", text="Valor indicado (qo)")
tabela.grid(row=4, column=0, columnspan=3, pady=20)     # Posicionamento

# Cria campos de entrada para adicionar dados à tabela
qi_label = tk.Label(janela, text="Valor verdadeiro (qi):")
qi_label.grid(row=1, column=0, padx=10, pady=5)
qi_entry = tk.Entry(janela)
qi_entry.grid(row=1, column=1, padx=10, pady=5)

qo_label = tk.Label(janela, text="Valor indicado (qo):")
qo_label.grid(row=2, column=0, padx=10, pady=5)
qo_entry = tk.Entry(janela)
qo_entry.grid(row=2, column=1, padx=10, pady=5)

entrada_label = tk.Label(janela, text="Valor indicado:")
entrada_label.grid(row=11, column=0, padx=10, pady=50)
entrada_entry = tk.Entry(janela)
entrada_entry.grid(row=11, column=1, padx=10, pady=50)


# Botão para importar um arquivo Excel
importar_excel_botao = tk.Button(janela, text="Importar dados", command=importar_excel)
importar_excel_botao.grid(row=0, column=2, padx=10, pady=5)

# Botão para adicionar um item à tabela
adicionar_botao = tk.Button(janela, text="Adicionar valores", command=adicionar_item)
adicionar_botao.grid(row=1, column=2, padx=10, pady=5)

# Botão para cálculos
calcular_botao = tk.Button(janela, text="   Equação da reta   ", command=calculos)
calcular_botao.grid(row=6, column=1)

# Botão para gerar curva de calibração
grafico_botao = tk.Button(janela, text="Curva de calibração", command=grafico)
grafico_botao.grid(row=10, column=1, pady=5)

# Botão para calcular valor verdadeiro
calcular_saida_botao = tk.Button(janela, text="Calcular valor real", command=saida)
calcular_saida_botao.grid(row=11, column=2, pady=5)

# Rótulo para mostrar o número de itens adicionados
N = tk.Label(janela, text=f"Quantidade: {contador_itens}")
N.grid(row=2, column=2)

# Rótulo para mostrar o valor de m
m_label = tk.Label(janela, text="m = ")
m_label.grid(row=7, column=1)

# Rótulo para mostrar o valor de b
b_label = tk.Label(janela, text="b = ")
b_label.grid(row=8, column=1)

# Rótulo para mostrar a equação da reta
equation_label = tk.Label(janela, text="qi = m.qo + b")
equation_label.grid(row=9, column=1)

# Rótulo para mostrar o valor verdadeiro de acordo com o valor indicado
saida_label = tk.Label(janela, text="")
saida_label.grid(row=11, column=3)

# Rótulo para mostrar erros de bias
bias_label = tk.Label(janela, text="")
bias_label.grid(row=12, column=3)

# Rótulo para mostrar imprecisão
imprecisao_label = tk.Label(janela, text="")
imprecisao_label.grid(row=13, column=3)



# Inicia o loop principal da interface gráfica
janela.mainloop()