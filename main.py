# Importação das bibliotecas necessárias
import math                      # Biblioteca para operações matemáticas (sqrt, ceil, etc.)
import pandas as pd              # Biblioteca para leitura e manipulação de dados (Excel)
import customtkinter as ctk     # Versão moderna do Tkinter para criação de interfaces
from tkinter import filedialog, messagebox  # Ferramentas para abrir arquivos e mostrar mensagens

# Configuração visual da interface
ctk.set_appearance_mode("System")       # Usa o modo claro/escuro do sistema operacional
ctk.set_default_color_theme("blue")     # Define o tema de cores padrão

# Classe principal da aplicação
class AppEstatistica(ctk.CTk):

    def __init__(self):
        """
        Construtor da aplicação.
        Inicializa a janela principal e chama o método que cria a interface.
        """
        super().__init__()
        self.title("Analisador de Dados Estatísticos")   # Título da janela
        self.geometry("1050x800")                        # Tamanho da janela
        self._configurar_interface()                     # Criação dos elementos visuais

    def _configurar_interface(self):
        """
        Cria todos os elementos gráficos da interface.
        """

        # Título da aplicação
        self.label_titulo = ctk.CTkLabel(
            self,
            text="Análise de Distribuição de Frequência",
            font=("Roboto", 24, "bold")
        )
        self.label_titulo.pack(pady=20)

        # Botão para importar um arquivo Excel
        self.btn_importar = ctk.CTkButton(
            self,
            text="Selecionar Arquivo Excel",
            command=self.processar_dados   # função chamada ao clicar
        )
        self.btn_importar.pack(pady=10)

        # Frame para exibir as estatísticas calculadas
        self.frame_stats = ctk.CTkFrame(self)
        self.frame_stats.pack(pady=10, padx=20, fill="x")

        # Label onde aparecerão média, mediana, etc.
        self.label_stats = ctk.CTkLabel(
            self.frame_stats,
            text="Aguardando importação...",
            font=("Roboto", 14),
            justify="left"
        )
        self.label_stats.pack(pady=15)

        # Frame rolável para mostrar a tabela de frequência
        self.frame_tabela = ctk.CTkScrollableFrame(self, width=1000, height=350)
        self.frame_tabela.pack(pady=10, padx=20)

        # Botão para limpar a tela
        self.btn_limpar = ctk.CTkButton(
            self,
            text="Limpar",
            fg_color="#d9534f",
            hover_color="#c9302c",
            command=self.limpar_tela
        )
        self.btn_limpar.pack(pady=10)

    def criar_cabecalho(self):
        """
        Cria o cabeçalho da tabela de distribuição de frequência.
        """

        # Nome das colunas da tabela
        colunas = [
            "Classe",        # Intervalo da classe
            "fi",            # Frequência simples
            "xm",            # Ponto médio da classe
            "xm * fi",       # Produto do ponto médio pela frequência
            "fa",            # Frequência acumulada
            "fr (%)",        # Frequência relativa (%)
            "fra (%)",       # Frequência relativa acumulada (%)
            "xm - media"     # Diferença entre ponto médio e média
        ]

        # Cria cada coluna visualmente
        for i, col in enumerate(colunas):
            lbl = ctk.CTkLabel(
                self.frame_tabela,
                text=col,
                font=("Roboto", 12, "bold"),
                fg_color="#3b8ed0",
                text_color="white",
                corner_radius=0
            )
            lbl.grid(row=0, column=i, sticky="nsew", padx=1, pady=1)

        # Faz as colunas ocuparem o espaço proporcionalmente
        for i in range(len(colunas)):
            self.frame_tabela.grid_columnconfigure(i, weight=1)

    def limpar_tela(self):
        """
        Remove todos os elementos da tabela e reinicia o texto de estatísticas.
        """
        for widget in self.frame_tabela.winfo_children():
            widget.destroy()

        self.label_stats.configure(text="Aguardando importação...")

    # -----------------------------
    # FUNÇÕES DE CÁLCULO ESTATÍSTICO
    # -----------------------------

    def calcular_classes(self, dados):
        """
        Cria os intervalos de classes da distribuição de frequência.

        dados -> lista com os valores importados do Excel

        retorna -> Dicionario com as divisões de classe contendo limite inferior, limite superior,
        frequência simples, frequência acumulada, ponto médio da classe, um indicador se é a ultima classe,
        além da amplitude da classe.
        """

        n = len(dados)               # número total de dados
        k = round(math.sqrt(n))      # número de classes (regra da raiz de n)
        h = math.ceil((max(dados) - min(dados)) / k)  # amplitude da classe

        limite_inf_inicial = min(dados)  # primeiro limite inferior

        lista_classes = []  # lista onde serão armazenadas as classes
        acumulado = 0       # frequência acumulada

        # Criação de cada classe
        for i in range(k):

            # Limite inferior da classe
            l_inf = limite_inf_inicial + (i * h)

            # Limite superior da classe
            l_sup = max(dados) if i == k - 1 else l_inf + h

            # Frequência simples (quantos dados estão na classe)
            if i == k - 1:
                fi = sum(1 for x in dados if l_inf <= x <= l_sup)
            else:
                fi = sum(1 for x in dados if l_inf <= x < l_sup)

            # Atualiza frequência acumulada
            acumulado += fi

            # Guarda as informações da classe em uma Lista de Dicionários chamada lista_classes
            lista_classes.append({
                'lim_inf': l_inf,       # limite inferior
                'lim_sup': l_sup,       # limite superior
                'fi': fi,               # frequência simples
                'fa': acumulado,        # frequência acumulada
                'xm': (l_inf + l_sup) / 2,  # ponto médio da classe
                'is_last': (i == k - 1) # indica se é a última classe
            })

        return lista_classes, h

    def calcular_estatisticas(self, lista_classes, n, h):
        """
        Parâmetros de entrada -> Lista_classes, n( número de elementos), h(amplitude de cada classe)
        Calcula média, mediana, classe modal e desvio padrão.
        """

        # -----------------------------
        # MÉDIA AGRUPADA
        # -----------------------------
        # Soma de (ponto médio * frequência) para cada dicionario de lista_classes
        soma_xm_fi = sum(c['xm'] * c['fi'] for c in lista_classes)

        # Média = soma(xm * fi) / n
        media = soma_xm_fi / n

        # -----------------------------
        # MEDIANA AGRUPADA
        # -----------------------------
        #pos_med tem como o objetivo descobrir a posição da mediana
        if n % 2 == 0:
            pos_med = n / 2
        else:
            pos_med = (n + 1) / 2

        # encontra a classe onde está a mediana percorrendo cada elemento do lista_classes e verificando se a frequência
        # acumulada é maior ou igual a posição da mediana. Assim que encontra retorna o index
        c_med_idx = next(i for i, c in enumerate(lista_classes) if c['fa'] >= pos_med)

        # Pega o dicionário usando o index em questão
        c_med = lista_classes[c_med_idx]

        # frequência acumulada anterior
        fa_ant = lista_classes[c_med_idx-1]['fa'] if c_med_idx > 0 else 0

        # fórmula da mediana para dados agrupados
        mediana = c_med['lim_inf'] + ((pos_med - fa_ant) / c_med['fi']) * h

        # -----------------------------
        # CLASSES MODAIS
        # -----------------------------
        # Determina a maior frequência dentro da lista de classes 
        max_fi = max(c['fi'] for c in lista_classes)

        # adiciona em indices_modais
        indices_modais = [i for i, c in enumerate(lista_classes) if c['fi'] == max_fi]

        intervalos_modais = []

        for idx in indices_modais:
            c = lista_classes[idx]
            sep = "|-|" if c['is_last'] else "|-"

            intervalos_modais.append(f"{c['lim_inf']:.1f} {sep} {c['lim_sup']:.1f}")

        label_moda = "Classe Modal" if len(indices_modais) == 1 else "Classes Modais"
        str_modas = ", ".join(intervalos_modais)

        # -----------------------------
        # DESVIO PADRÃO
        # -----------------------------
        variancia = sum(c['fi'] * (c['xm'] - media)**2 for c in lista_classes) / n
        desvio_padrao = math.sqrt(variancia)

        return media, mediana, label_moda, str_modas, desvio_padrao

    def renderizar_tabela(self, lista_classes, n, media):
        """
        Mostra na interface a tabela de distribuição de frequência.
        """

        for i, c in enumerate(lista_classes):

            # frequência relativa
            fr = (c['fi'] / n) * 100

            # frequência relativa acumulada
            fra = (c['fa'] / n) * 100

            sep = "|-|" if c['is_last'] else "|-"

            valores = [
                f"{c['lim_inf']:.1f} {sep} {c['lim_sup']:.1f}",
                str(c['fi']),
                f"{c['xm']:.2f}",
                f"{c['xm']*c['fi']:.2f}",
                str(c['fa']),
                f"{fr:.1f}%",
                f"{fra:.1f}%",
                f"{c['xm']-media:.2f}"
            ]

            cor_fundo = "#FFFFFF"

            for col_idx, valor in enumerate(valores):
                lbl = ctk.CTkLabel(
                    self.frame_tabela,
                    text=valor,
                    fg_color=cor_fundo,
                    corner_radius=0
                )

                lbl.grid(row=i+1, column=col_idx, sticky="nsew", padx=1, pady=1)

    # -----------------------------
    # FUNÇÃO PRINCIPAL
    # -----------------------------

    def processar_dados(self):
        """
        Função principal da aplicação.

        Fluxo:
        1) Abre o Excel
        2) Lê os dados
        3) Calcula as classes
        4) Calcula estatísticas
        5) Mostra resultados na interface
        """

        caminho = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )

        if not caminho:
            return

        try:

            # Limpa a interface antes de carregar novos dados
            self.limpar_tela()
            self.criar_cabecalho()

            # Leitura do Excel
            df = pd.read_excel(caminho, header=None)

            # Pega a primeira coluna
            dados = df.iloc[:, 0].dropna().astype(float).tolist()

            dados.sort()

            n = len(dados)

            if n == 0:
                messagebox.showwarning("Aviso", "O arquivo está vazio.")
                return

            # Cálculos estatísticos
            lista_classes, h = self.calcular_classes(dados)

            media, mediana, lbl_mod, txt_mod, desvio = self.calcular_estatisticas(
                lista_classes, n, h
            )

            # Mostra tabela
            self.renderizar_tabela(lista_classes, n, media)

            # Texto com as estatísticas finais
            texto_stats = (
                f"N: {n}  |  Média (x̄): {media:.2f}  |  Mediana: {mediana:.2f}\n"
                f"{lbl_mod}: {txt_mod}  |  Desvio Padrão Populacional (σ): {desvio:.2f}"
            )

            self.label_stats.configure(text=texto_stats)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento:\n{str(e)}")


# Executa o programa
if __name__ == "__main__":
    app = AppEstatistica()
    app.mainloop()