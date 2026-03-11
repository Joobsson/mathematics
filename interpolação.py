import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import warnings
import tkinter as tk
from tkinter import filedialog
import sys
import os

# Tratamento para retrocompatibilidade do NumPy
try:
    from numpy.exceptions import RankWarning
except ImportError:
    RankWarning = np.RankWarning

def selecionar_arquivo_dinamicamente():
    root = tk.Tk()
    root.withdraw() 
    caminho = filedialog.askopenfilename(
        title="Selecione a Planilha de Metadados Espectrais",
        filetypes=[("Arquivos CSV", "*.csv"), ("Arquivos Excel", "*.xlsx *.xls")]
    )
    return caminho

def forcar_polinomio_grau_maximo(caminho_arquivo: str, nome_aba: str):
    """
    Força a modelagem de um polinômio global de grau N-1 para interpolação exata.
    """
    print(f"Lendo base matricial em: {caminho_arquivo}")
    
    try:
        if caminho_arquivo.lower().endswith('.csv'):
            df = pd.read_csv(caminho_arquivo)
        else:
            df = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)
    except Exception as e:
        print(f"Falha catastrófica de leitura: {e}")
        return None

    col_x = df.columns[0]
    col_y = df.columns[1]
    
    x = df[col_x].values
    y = df[col_y].values

    n_pontos = len(x)
    grau_maximo = n_pontos - 1  # Limite teórico máximo para N pontos distintos
    
    print(f"\n--- Iniciando Cômputo de Matriz Mal Condicionada ---")
    print(f"Total de pontos na amostragem (N): {n_pontos}")
    print(f"Grau do Polinômio Global estabelecido: {grau_maximo}")
    
    # A supressão deste aviso é imperativa, pois o NumPy gritará sobre a singularidade da matriz
    warnings.simplefilter('ignore', RankWarning)

    # Para evitar overflow imediato durante o cômputo de x^942, 
    # é necessário normalizar o domínio das frequências para o intervalo [-1, 1].
    # Sem esta normalização espacial, a CPU abortaria o cálculo por estouro de limite (NaN).
    x_mean = np.mean(x)
    x_std = np.std(x)
    x_normalizado = (x - x_mean) / x_std

    print("Calculando coeficientes. Isso pode exigir grande alocação de registos na FPU...")
    coeficientes = np.polyfit(x_normalizado, y, deg=grau_maximo)
    polinomio = np.poly1d(coeficientes)
    
    print("Cálculo concluído com sucesso analítico (sujeito à truncagem de ponto flutuante).")

    # ==========================================
    # Rotina de Exportação Algébrica
    # ==========================================
    diretorio_base = os.path.dirname(caminho_arquivo)
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    caminho_txt = os.path.join(diretorio_base, f"{nome_base}_Polinomio_Grau_{grau_maximo}.txt")
    caminho_csv = os.path.join(diretorio_base, f"{nome_base}_Coeficientes_Globais.csv")

    # Exportação Matricial (.csv)
    df_coefs = pd.DataFrame({
        'Grau_Termo': range(grau_maximo, -1, -1),
        'Coeficiente_a_j': coeficientes
    })
    df_coefs.to_csv(caminho_csv, index=False, sep=';', decimal=',')

    # Exportação Textual Analítica (.txt)
    with open(caminho_txt, 'w', encoding='utf-8') as f:
        f.write(f"--- Polinômio de Interpolação Global de Grau {grau_maximo} ---\n")
        f.write(f"Atenção: A variável independente x' encontra-se normalizada.\n")
        f.write(f"Relação Espacial: x' = (Frequencia_Hz - {x_mean:.6f}) / {x_std:.6f}\n\n")
        f.write("P(x') = \n")
        for i, coef in enumerate(coeficientes):
            grau_atual = grau_maximo - i
            sinal = " + " if coef >= 0 else " - "
            valor_abs = abs(coef)
            f.write(f"{sinal}{valor_abs:.10e} * (x')^{grau_atual}\n")

    print(f"\n[SUCESSO] O polinômio colossal foi exportado:")
    print(f"Equação P(x): {caminho_txt}")
    print(f"Vetor de Coeficientes: {caminho_csv}")

    # ==========================================
    # Renderização Gráfica
    # ==========================================
    plt.figure(figsize=(12, 6))
    plt.scatter(x, y, color='blue', s=10, label='Sinal Original', zorder=3)
    
    # Plotagem do polinômio na base espacial normalizada, retornada ao domínio original
    x_denso = np.linspace(min(x), max(x), 5000)
    x_denso_norm = (x_denso - x_mean) / x_std
    y_fit = polinomio(x_denso_norm)
    
    # Restrição do eixo Y para evitar que o Fenômeno de Runge distorça completamente a visualização
    limite_y = max(y) * 1.5 
    plt.ylim(min(y) - limite_y*0.1, limite_y)
    
    plt.plot(x_denso, y_fit, color='red', linestyle='-', label=f'Polinômio Global (Grau {grau_maximo})', linewidth=1)
    
    plt.title(f'Ajuste Analítico Forçado de Polinômio Global (Grau Máximo: {grau_maximo})', fontsize=14)
    plt.xlabel(f'{col_x} (Hz)', fontsize=12)
    plt.ylabel(f'{col_y} (g)', fontsize=12)
    plt.grid(True, linestyle=':', alpha=0.8)
    plt.legend()
    plt.tight_layout()
    plt.show()

    return polinomio

if __name__ == "__main__":
    caminho = selecionar_arquivo_dinamicamente()
    if not caminho:
        sys.exit()
    forcar_polinomio_grau_maximo(caminho, 'Discretização Matricial')