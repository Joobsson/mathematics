import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import CubicSpline
import tkinter as tk
from tkinter import filedialog
import sys

def selecionar_ficheiro_dinamicamente():
    """
    Invoca a interface gráfica nativa do sistema operativo para a seleção do ficheiro de dados.
    Retorna o caminho absoluto do ficheiro selecionado.
    """
    root = tk.Tk()
    root.withdraw() # Oculta a janela principal para não sobrecarregar o ecrã
    
    caminho_absoluto = filedialog.askopenfilename(
        title="Selecione o Ficheiro de Metadados Espetrais",
        filetypes=[
            ("Ficheiros CSV", "*.csv"),
            ("Ficheiros Excel", "*.xlsx *.xls"),
            ("Todos os Ficheiros", "*.*")
        ]
    )
    return caminho_absoluto

def interpolar_espetro_spline(caminho_ficheiro: str, nome_aba: str):
    """
    Efetua a leitura dos dados e aplica a Interpolação por Splines Cúbicas 
    para mapear exatamente as componentes de vibração (amplitudes em função da frequência).
    """
    print(f"A iniciar a leitura do ficheiro: {caminho_ficheiro}")
    
    try:
        # Suporte dinâmico para leitura direta de ficheiros .csv ou .xlsx
        if caminho_ficheiro.lower().endswith('.csv'):
            df = pd.read_csv(caminho_ficheiro)
        else:
            df = pd.read_excel(caminho_ficheiro, sheet_name=nome_aba)
            
    except ValueError:
        print(f"Erro: A aba '{nome_aba}' não foi localizada no ficheiro.")
        return None
    except Exception as e:
        print(f"Erro crítico durante a extração de dados: {e}")
        return None

    # Extração das grandezas físicas assumindo a estrutura vetorial padrão do espetro
    col_x = df.columns[0] # Esperado: Frequencia_Hz
    col_y = df.columns[1] # Esperado: Amplitude_g
    
    x = df[col_x].values
    y = df[col_y].values

    # Condição de estabilidade numérica: o vetor de abscissas (frequências) deve ser estritamente crescente.
    # A ordenação previne singularidades na matriz tridiagonal durante a resolução do sistema linear da Spline.
    indices_ordenados = np.argsort(x)
    x = x[indices_ordenados]
    y = y[indices_ordenados]

    # Implementação do modelo de Interpolação Segmentada.
    # bc_type='natural' impõe que a segunda derivada nos extremos seja nula (S''(x_0) = S''(x_n) = 0).
    spline = CubicSpline(x, y, bc_type='natural')
    
    print(f"\n--- Processamento Matemático Concluído ---")
    print(f"Método Numérico: Splines Cúbicas (Classe C²)")
    print(f"Dimensão da base de dados interpolada: {len(x)} pontos")

    # Configuração do ambiente gráfico para validação da discretização
    plt.figure(figsize=(12, 6))
    
    # Representação do sinal discreto original
    plt.scatter(x, y, color='#1f77b4', s=15, label='Sinal Discreto Original (Amostragem)', zorder=3)
    
    # Adensamento do domínio de frequência para demonstrar a suavidade analítica da curva gerada
    x_denso = np.linspace(min(x), max(x), 5000)
    y_interp = spline(x_denso)
    
    # Traçado da função contínua interpoladora que converge de forma exata sobre os pontos azuis
    plt.plot(x_denso, y_interp, color='#d62728', linestyle='-', label='Interpolação Spline Cúbica', linewidth=1.5, zorder=2)
    
    plt.title('Análise Espetral: Interpolação Exata (Amplitude vs Frequência)', fontsize=14, fontweight='bold')
    plt.xlabel(f'{col_x}', fontsize=12)
    plt.ylabel(f'{col_y}', fontsize=12)
    plt.grid(True, linestyle=':', alpha=0.8)
    plt.legend()
    plt.tight_layout()
    plt.show()

    return spline

# ==========================================
# Rotina Principal de Execução
# ==========================================
if __name__ == "__main__":
    
    # Acionamento da interface gráfica para captura do diretório
    caminho = selecionar_ficheiro_dinamicamente()
    
    # Validação da operação de seleção
    if not caminho:
        print("Operação cancelada pelo utilizador. A rotina será terminada.")
        sys.exit()
        
    aba_alvo = 'Discretização Matricial'
    
    # Execução da sub-rotina de interpolação avançada
    funcao_interpoladora = interpolar_espetro_spline(
        caminho_ficheiro=caminho, 
        nome_aba=aba_alvo
    )