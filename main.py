import math
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class AppEstatistica(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Analisador de Dados Estatísticos")
        self.geometry("1050x800")
        self._configurar_interface()

    def _configurar_interface(self):
        """Inicializa os elementos visuais da janela."""
        self.label_titulo = ctk.CTkLabel(self, text="Análise de Distribuição de Frequência", font=("Roboto", 24, "bold"))
        self.label_titulo.pack(pady=20)

        self.btn_importar = ctk.CTkButton(self, text="Selecionar Arquivo Excel", command=self.processar_dados)
        self.btn_importar.pack(pady=10)

        self.frame_stats = ctk.CTkFrame(self)
        self.frame_stats.pack(pady=10, padx=20, fill="x")
        self.label_stats = ctk.CTkLabel(self.frame_stats, text="Aguardando importação...", font=("Roboto", 14), justify="left")
        self.label_stats.pack(pady=15)

        self.frame_tabela = ctk.CTkScrollableFrame(self, width=1000, height=350)
        self.frame_tabela.pack(pady=10, padx=20)

        self.btn_limpar = ctk.CTkButton(self, text="Limpar", fg_color="#d9534f", hover_color="#c9302c", command=self.limpar_tela)
        self.btn_limpar.pack(pady=10)

    def criar_cabecalho(self):
        colunas = ["Classe", "fi", "xm", "xm * fi", "fa", "fr (%)", "fra (%)", "xm - media"]
        for i, col in enumerate(colunas):
            lbl = ctk.CTkLabel(self.frame_tabela, text=col, font=("Roboto", 12, "bold"), 
                               fg_color="#3b8ed0", text_color="white", corner_radius=0)
            lbl.grid(row=0, column=i, sticky="nsew", padx=1, pady=1)
        for i in range(len(colunas)):
            self.frame_tabela.grid_columnconfigure(i, weight=1)

    def limpar_tela(self):
        for widget in self.frame_tabela.winfo_children():
            widget.destroy()
        self.label_stats.configure(text="Aguardando importação...")

    # --- FUNÇÕES DE CÁLCULO (Lógica Separada) ---

    def calcular_classes(self, dados):
        """Define intervalos, frequências e pontos médios."""
        n = len(dados)
        k = round(math.sqrt(n))
        h = math.ceil((max(dados) - min(dados)) / k)
        limite_inf_inicial = min(dados)
        
        lista_classes = []
        acumulado = 0
        
        for i in range(k):
            l_inf = limite_inf_inicial + (i * h)
            l_sup = max(dados) if i == k - 1 else l_inf + h
            
            # Frequência simples (fi)
            if i == k - 1:
                fi = sum(1 for x in dados if l_inf <= x <= l_sup)
            else:
                fi = sum(1 for x in dados if l_inf <= x < l_sup)
            
            acumulado += fi
            lista_classes.append({
                #fi é a frequência simples, fa é a frequência acumulada, xm é o ponto médio da classe
                'lim_inf': l_inf, 
                'lim_sup': l_sup,
                'fi': fi,
                'fa': acumulado, 
                'xm': (l_inf + l_sup) / 2,
                'is_last': (i == k - 1)
            })
        return lista_classes, h

    def calcular_estatisticas(self, lista_classes, n, h):
        """Calcula média, mediana, modas e desvio padrão das classes."""
        # 1. Média Agrupada
        # Multiplica a média da classe pela frequencia simples e soma para todas as classes, depois divide pelo total de dados
        soma_xm_fi = sum(c['xm'] * c['fi'] for c in lista_classes)
        media = soma_xm_fi / n

        # 2. Mediana Agrupada
        if n % 2 == 0:
            pos_med = n / 2
        else:
            pos_med = (n + 1) / 2
        
        #next é usado aqui para identificar a primeira classe cuja frequência acumulada é maior ou igual à posição da mediana
        c_med_idx = next(i for i, c in enumerate(lista_classes) if c['fa'] >= pos_med)
        c_med = lista_classes[c_med_idx]
        fa_ant = lista_classes[c_med_idx-1]['fa'] if c_med_idx > 0 else 0
        
        mediana = c_med['lim_inf'] + ((pos_med - fa_ant) / c_med['fi']) * h

        # 3. Classes Modais
        #pega o valor da maior frequência simples
        max_fi = max(c['fi'] for c in lista_classes)
        # pega os indices de todas as classes que possuem essa frequência simples máxima
        indices_modais = [i for i, c in enumerate(lista_classes) if c['fi'] == max_fi]
        
        intervalos_modais = []
        for idx in indices_modais:
            c = lista_classes[idx]
            sep = "|-|" if c['is_last'] else "|-"
            #pega os intervalos modais
            intervalos_modais.append(f"{c['lim_inf']:.1f} {sep} {c['lim_sup']:.1f}")
        
        label_moda = "Classe Modal" if len(indices_modais) == 1 else "Classes Modais"
        str_modas = ", ".join(intervalos_modais)

        # 4. Desvio Padrão Populacional
        soma_var = sum((c['xm'] - media) for c in lista_classes)
        desvio_padrao = math.sqrt(soma_var / n)

        return media, mediana, label_moda, str_modas, desvio_padrao

    def renderizar_tabela(self, lista_classes, n, media):
        """Desenha a tabela na interface."""
        for i, c in enumerate(lista_classes):
            fr = (c['fi'] / n) * 100
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

            cor_fundo = "#2b2b2b" if i % 2 == 0 else "#333333"
            for col_idx, valor in enumerate(valores):
                lbl = ctk.CTkLabel(self.frame_tabela, text=valor, fg_color=cor_fundo, corner_radius=0)
                lbl.grid(row=i+1, column=col_idx, sticky="nsew", padx=1, pady=1)

    # --- FUNÇÃO PRINCIPAL (Orquestradora) ---

    def processar_dados(self):
        caminho = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx *.xls")])
        if not caminho: return

        try:
            # Preparação inicial
            self.limpar_tela()
            self.criar_cabecalho()
            
            # Carga dos dados
            df = pd.read_excel(caminho, header=None)
            dados = df.iloc[:, 0].dropna().astype(float).tolist()
            dados.sort()
            n = len(dados)

            if n == 0:
                messagebox.showwarning("Aviso", "O arquivo está vazio.")
                return

            # Execução da lógica separada
            lista_classes, h = self.calcular_classes(dados)
            media, mediana, lbl_mod, txt_mod, desvio = self.calcular_estatisticas(lista_classes, n, h)
            
            # Interface
            self.renderizar_tabela(lista_classes, n, media)
            
            texto_stats = (
                f"N: {n}  |  Média (x̄): {media:.2f}  |  Mediana: {mediana:.2f}\n"
                f"{lbl_mod}: {txt_mod}  |  Desvio Padrão Populacional (σ): {desvio:.2f}"
            )
            self.label_stats.configure(text=texto_stats)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento:\n{str(e)}")

if __name__ == "__main__":
    app = AppEstatistica()
    app.mainloop()