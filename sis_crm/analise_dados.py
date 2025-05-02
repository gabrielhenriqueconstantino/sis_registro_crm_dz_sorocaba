import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import sqlite3
from datetime import datetime
import requests  # Para integração com a API de CEP
import re
from PIL import Image, ImageTk
import os
import matplotlib.pyplot as plt
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import Counter
import matplotlib.cm as cm
import numpy as np
from matplotlib import rcParams

from pathlib import Path

class JanelaAnalise(ctk.CTkToplevel):
    def __init__(self, master=None):
        super().__init__(master)
        self.title("Análise de Dados")
        
        # Inicialmente invisível
        self.attributes("-alpha", 0.0)
        
        # Ícone da janela (mesmo código anterior)
        caminho_icone = Path(__file__).parent / "img" / "icons" / "icons_window" / "icon_dados.ico"
        try:
            if os.name == 'nt':
                self.after(200, lambda: self.iconbitmap(caminho_icone))
            else:
                img = Image.open(caminho_icone)
                icon = ctk.CTkImage(img)
                self.after(200, lambda: self.wm_iconphoto(True, icon._photo_image))
        except Exception as e:
            print(f"Icone não encontrado ou não suportado: {e}")
        
        # Tamanho e centralização
        largura_janela = 1200
        altura_janela = 600
        largura_tela = self.winfo_screenwidth()
        altura_tela = self.winfo_screenheight()
        x = (largura_tela // 2) - (largura_janela // 2)
        y = (altura_tela // 2) - (altura_janela // 2)
        self.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

        # Garantir foco e topo
        self.attributes('-topmost', True)
        self.after(100, lambda: self.attributes('-topmost', False))
        self.lift()
        self.focus_force()
        self.grab_set()

        # Fonte padrão
        self.fonte_padrao = ctk.CTkFont(family="Poppins", size=12)

        # Animação suave
        self.after(0, self.animar_aparicao)

    def animar_aparicao(self):
        """Anima a opacidade da janela para um efeito suave de abertura."""
        def fade_in(alpha=0.0):
            if alpha < 1.0:
                alpha += 0.05
                self.attributes("-alpha", alpha)
                self.after(15, lambda: fade_in(alpha))
            else:
                self.attributes("-alpha", 1.0)
        fade_in()

        # Fonte padrão
        self.fonte_padrao = ctk.CTkFont(family="Poppins", size=12)

        # Estilo para o DateEntry (só criar uma vez)
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Custom.DateEntry",
                    fieldbackground="#2b2b2b",
                    background="#2b2b2b",
                    foreground="white",
                    arrowcolor="white",
                    bordercolor="#3a3a3a",
                    lightcolor="#3a3a3a",
                    darkcolor="#3a3a3a")

        # Frame principal
        frame_analise = ctk.CTkFrame((self), width=760, height=560, corner_radius=15, border_color="white")
        frame_analise.pack(padx=20, pady=20, fill="both", expand=True)

        # Frame lateral
        container_lateral = ctk.CTkFrame(frame_analise)
        container_lateral.pack(fill="both", expand=True)

        # Container datas
        container_datas = ctk.CTkFrame(container_lateral, corner_radius=10, border_width=2, border_color="white")
        container_datas.pack(side="left", padx=10, pady=10, anchor="n")

        # Intervalo dos dados
        label_intervalo = ctk.CTkLabel(container_datas,
                                   text="Intervalo dos dados",
                                   font=ctk.CTkFont("Poppins", size=14, weight="bold"),
                                   anchor="w")
        label_intervalo.grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(10, 0))

        # Campo "DE" com estilo escuro
        self.data_de = DateEntry(container_datas, date_pattern='dd/mm/yyyy',
                             style="Custom.DateEntry",
                             borderwidth=1,
                             font=("Arial", 10),
                             justify="center",
                             year=2024, month=10, day=10,
                             locale="pt_BR")
        self.data_de.grid(row=1, column=0, padx=(10, 5), pady=10)

        # Texto "até"
        label_ate = ctk.CTkLabel(container_datas,
                             text="até",
                             font=ctk.CTkFont(size=12),
                             anchor="center")
        label_ate.grid(row=1, column=1, padx=5, pady=10)

        # Campo "ATÉ" com estilo escuro também
        self.data_ate = DateEntry(container_datas, date_pattern='dd/mm/yyyy',
                              style="Custom.DateEntry",
                              borderwidth=1,
                              font=("Arial", 10),
                              justify="center",
                              locale="pt_BR")
        self.data_ate.grid(row=1, column=2, padx=(5, 10), pady=10)

        # Botão para aplicar filtro
        botao_filtrar = ctk.CTkButton(container_datas,
                          text="APLICAR",
                          font=("Poppins", 14, "bold"),
                          fg_color="#0078D7",
                          hover_color="#005A9E",
                          border_width=1,
                          border_color="white",
                          text_color="white",
                          command=lambda: [self.aplicar_filtro(self.frame_exclusivo_grafico), 
                                          self.atualizar_resultados(frame_resultados)])
        botao_filtrar.grid(row=2, column=0, columnspan=3, pady=(0, 10), padx=10, sticky="ew")

        # Container de resultados gerais
        frame_resultados = ctk.CTkFrame(container_datas, corner_radius=10, border_width=2, border_color="white")
        frame_resultados.grid(row=3, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")

        # Título dos resultados
        label_resultados = ctk.CTkLabel(frame_resultados,
                         text="Estatísticas Gerais",
                         font=ctk.CTkFont("Poppins", size=14, weight="bold"),
                         anchor="w")
        label_resultados.pack(padx=10, pady=(10, 5), anchor="w")
            
        # Labels para as estatísticas (serão atualizadas)
        self.label_total = ctk.CTkLabel(frame_resultados, text="Total do período: ", anchor="w", font=self.fonte_padrao)
        self.label_total.pack(padx=10, pady=2, fill="x")
    
        self.label_area_mais = ctk.CTkLabel(frame_resultados, text="Área mais requisitada: ", anchor="w", font=self.fonte_padrao)
        self.label_area_mais.pack(padx=10, pady=2, fill="x")
    
        self.label_media_semanal = ctk.CTkLabel(frame_resultados, text="Média semanal: ", anchor="w", font=self.fonte_padrao)
        self.label_media_semanal.pack(padx=10, pady=2, fill="x")
    
        self.label_media_diaria = ctk.CTkLabel(frame_resultados, text="Média diária: ", anchor="w", font=self.fonte_padrao)
        self.label_media_diaria.pack(padx=10, pady=(2, 10), fill="x")

        #Container do ranking
        container_ranking = ctk.CTkFrame(container_datas, corner_radius=10, border_width=2, border_color="white")
        container_ranking.grid(row=4, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")

        # Label "Ranking"
        label_ranking = ctk.CTkLabel(container_ranking,
                                text="Ranking: Bairros mais requisitados",
                                font=ctk.CTkFont("Poppins", size=14, weight="bold"),
                                anchor="w")
        label_ranking.pack(padx=10, pady=(10, 5), anchor="w")

        # Frame para os itens do ranking
        frame_itens_ranking = ctk.CTkFrame(container_ranking, fg_color="transparent")
        frame_itens_ranking.pack(padx=10, pady=(0, 10), fill="both", expand=True)

        # Limpar a lista de labels do ranking
        self.labels_ranking = []

        # Criar as 5 labels do ranking
        for i in range(1, 6):
            frame_item = ctk.CTkFrame(frame_itens_ranking, fg_color="transparent")
            frame_item.pack(fill="x", pady=1)

            # Bullet point (∙)
            label_bullet = ctk.CTkLabel(frame_item, text="∙", width=10, anchor="w")
            label_bullet.pack(side="left")

            # Posição (1º, 2º, etc.)
            label_posicao = ctk.CTkLabel(frame_item, text=f"{i}º - ", width=30, anchor="w")
            label_posicao.pack(side="left")

            # Bairro e quantidade (será atualizado)
            label_conteudo = ctk.CTkLabel(frame_item, text="", anchor="w", font=self.fonte_padrao)
            label_conteudo.pack(side="left", fill="x", expand=True)

            self.labels_ranking.append((label_posicao, label_conteudo))

        # Botão "Ver tudo"
        botao_ver_tudo = ctk.CTkButton(
        container_ranking,
        text="VER TUDO",
        font=("Poppins", 14, "bold"),
        text_color="white",
        fg_color="#0078D7",
        hover_color="#005A9E",
        border_width=1,
        border_color="white",
        command=self.mostrar_todos_bairros
        )
        botao_ver_tudo.pack(pady=(0, 10), padx=10, fill="x")

        # Container para os botões de ação - com expansão vertical
        container_botoes_finais = ctk.CTkFrame(
        container_datas, 
        corner_radius=10, 
        border_width=2, 
        border_color="white",
        height=70  # Altura fixa para garantir espaço
        )
        container_botoes_finais.grid(row=5, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="nsew")
        container_botoes_finais.grid_propagate(False)  # Impede que o frame redimensione automaticamente
        container_botoes_finais.pack_propagate(False)  # Garante que mantém o tamanho definido

        # Configurar grid para os botões com padding interno
        container_botoes_finais.grid_columnconfigure(0, weight=1)
        container_botoes_finais.grid_columnconfigure(1, weight=1)
        container_botoes_finais.grid_rowconfigure(0, weight=1)

        # Botão "❎" (fechar janela) - com margem interna
        self.icone_fechar = ctk.CTkImage(Image.open("img/icons/icons_app/close.png"), size=(30, 30))
        botao_fechar = ctk.CTkButton(
        container_botoes_finais,
        image=self.icone_fechar,
        text="",
        width=52,
        height=52,
        font=ctk.CTkFont(size=23),
        fg_color="#FF5555",
        hover_color="#FF3333",
        corner_radius=10,
        border_width=1,
        border_color="white",
        command=self.fechar_janela_analise
        )
        botao_fechar.grid(row=0, column=0, padx=(0, 5), pady=8, sticky="e")  # Adicionado pady=8

        self.icone_sol = ctk.CTkImage(Image.open("img/icons/icons_app/sun.png"), size=(30, 30))
        self.icone_lua = ctk.CTkImage(Image.open("img/icons/icons_app/moon.png"), size=(30, 30))
        self.icone_inicial = self.icone_sol if ctk.get_appearance_mode() == "Dark" else self.icone_lua
        self.botao_tema = ctk.CTkButton(
        container_botoes_finais,
        image=self.icone_inicial,
        text="",
        width=52,
        height=52,
        font=ctk.CTkFont(size=24),
        fg_color="#4B5563",
        hover_color="#374151",
        corner_radius=10,
        border_width=1,
        border_color="white",
        command=self.alternar_tema
        )
        self.botao_tema.grid(row=0, column=1, padx=(5, 0), pady=8, sticky="w")  # Adicionado pady=8

        # Container à direita (gráfico)
        container_grafico = ctk.CTkFrame(container_lateral, corner_radius=10, border_width=2, border_color="white")
        container_grafico.pack(side="left", padx=10, pady=10, fill="both", expand=True)

        # Frame para os controles (checkboxes)
        frame_checkboxes = ctk.CTkFrame(container_grafico, fg_color="transparent")
        frame_checkboxes.pack(fill="x", padx=10, pady=10)

        # Primeiro container (Tipos de Gráfico)
        container_tipo_grafico = ctk.CTkFrame(frame_checkboxes, 
                                    corner_radius=8, 
                                    border_width=1, 
                                    border_color="white")
        container_tipo_grafico.pack_propagate(False)
        container_tipo_grafico.pack(side="left", padx=(0, 10))
        container_tipo_grafico.configure(width=150, height=80)

        self.checkbox_barras = ctk.CTkCheckBox(container_tipo_grafico, 
            text="Barras",
            fg_color="#0078D7",
            hover_color="#005A9E",
            font=self.fonte_padrao,
            command=lambda: self.mudar_visualizacao('barras'))
        self.checkbox_barras.pack(pady=(10, 0), padx=10, anchor="w")
        self.checkbox_barras.select()

        self.checkbox_pizza = ctk.CTkCheckBox(container_tipo_grafico, 
            text="Pizza",
            fg_color="#0078D7",
            hover_color="#005A9E",
            font=self.fonte_padrao,
            command=lambda: self.mudar_visualizacao('pizza'))
        self.checkbox_pizza.pack(pady=(5, 10), padx=10, anchor="w")
        
        # Segundo container (Modo de Agrupamento)
        container_agrupamento = ctk.CTkFrame(frame_checkboxes, 
                                    corner_radius=8, 
                                    border_width=1, 
                                    border_color="white")
        container_agrupamento.pack_propagate(False)
        container_agrupamento.pack(side="left", padx=(0, 10))
        container_agrupamento.configure(width=200, height=80)

        self.checkbox_quantidade = ctk.CTkCheckBox(container_agrupamento, 
                                     text="Por quantidade",
                                     fg_color="#0078D7",
                                     hover_color="#005A9E",
                                     font=self.fonte_padrao,
                                     command=lambda: [self.checkbox_assunto.deselect(), 
                                    self.alternar_modo_visualizacao()])                      
        self.checkbox_quantidade.pack(pady=(10, 0), padx=10, anchor="w")
        self.checkbox_quantidade.select()
        
        self.checkbox_assunto = ctk.CTkCheckBox(container_agrupamento, 
                          text="Por assunto",
                          fg_color="#0078D7",
                          hover_color="#005A9E",
                          font=self.fonte_padrao,
                          command=lambda: [self.checkbox_quantidade.deselect(), 
                                          self.alternar_modo_visualizacao()])
        self.checkbox_assunto.pack(pady=(5, 10), padx=10, anchor="w")

        # Terceiro container (Áreas Municipais) - agora armazenamos como atributo
        self.container_areas = ctk.CTkFrame(frame_checkboxes, 
                         corner_radius=8, 
                         border_width=1, 
                         border_color="white")
        self.container_areas.pack_propagate(False)
        self.container_areas.pack(side="left", fill="x", expand=True)
        self.container_areas.configure(height=80)

        # Frame interno para as checkboxes das áreas
        frame_areas = ctk.CTkFrame(self.container_areas, fg_color="transparent")
        frame_areas.pack(padx=10, pady=10, fill="both", expand=True)

        
        # Frame exclusivo para o gráfico
        self.frame_exclusivo_grafico = ctk.CTkFrame(container_grafico, 
                                     border_width=2, 
                                     border_color="white",
                                     height=300)
        self.frame_exclusivo_grafico.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.frame_exclusivo_grafico.pack_propagate(False)

        # Lista de áreas municipais
        self.areas = ["Sudoeste", "Noroeste", "Centro-Sul", "Centro-Norte", "Norte", "Leste"]
        self.checkbox_areas = []

        # Configuração do grid
        LINHAS = 2
        COLUNAS = 3
        PADX = 15  # Espaçamento horizontal entre checkboxes
        PADY = 4  # Espaçamento vertical entre checkboxes

        # Adiciona as checkboxes em uma grade 2x3
        for i, area in enumerate(self.areas):
            linha = i // COLUNAS
            coluna = i % COLUNAS
    
            checkbox = ctk.CTkCheckBox(frame_areas, 
                  text=area,
                  font=self.fonte_padrao,
                  fg_color="#0078D7",
                  hover_color="#005A9E",
                  command=lambda: self.atualizar_grafico_por_checkbox(self.frame_exclusivo_grafico))
            checkbox.grid(row=linha, column=coluna, padx=PADX, pady=PADY, sticky="nswe")
            checkbox.select()
            self.checkbox_areas.append(checkbox)

        # Configura o peso das colunas para centralizar
        for col in range(COLUNAS):
            frame_areas.grid_columnconfigure(col, weight=1)

        # Criar gráfico inicial
        self.grafico_barras(self.frame_exclusivo_grafico)

    def aplicar_filtro(self, frame_grafico):
        """Aplica o filtro de data e atualiza o gráfico"""
        data_inicio = self.data_de.get_date()
        data_fim = self.data_ate.get_date()
        
        # Verifica se a data de início é menor que a data de fim
        if data_inicio > data_fim:
            messagebox.showerror("Erro", "A data de início não pode ser maior que a data de fim!")
            return
        
        # Verifica qual tipo de gráfico está selecionado e atualiza
        if self.checkbox_barras.get():
            self.grafico_barras(frame_grafico)
        else:
            self.atualizar_grafico_pizza(frame_grafico)

    def alternar_modo_visualizacao(self):
        """Alterna entre visualização por área e por assunto"""
        try:
            if self.checkbox_assunto.get():
                # Desmarca o checkbox "Por quantidade"
                self.checkbox_quantidade.deselect()
            
                # Esconder o container de áreas
                self.container_areas.pack_forget()

                # Atualizar o gráfico para mostrar por assunto
                if self.checkbox_barras.get():
                    self.grafico_barras_assuntos(self.frame_exclusivo_grafico)
                else:
                    self.atualizar_grafico_pizza_assuntos(self.frame_exclusivo_grafico)
            else:
                # Marca o checkbox "Por quantidade"
                self.checkbox_quantidade.select()
                
                # Mostrar o container de áreas
                self.container_areas.pack(side="left", fill="x", expand=True)

                # Atualizar o gráfico para mostrar por área
                if self.checkbox_barras.get():
                    self.grafico_barras(self.frame_exclusivo_grafico)
                else:
                    self.atualizar_grafico_pizza(self.frame_exclusivo_grafico)
                
        except Exception as e:
            print(f"Erro ao alternar modo de visualização: {e}")

    def mudar_visualizacao(self, tipo):
        """Muda entre visualização em barras e pizza"""
        try:
            if tipo == 'barras':
                self.checkbox_pizza.deselect()
                self.checkbox_barras.select()
                if self.checkbox_assunto.get():
                    self.grafico_barras_assuntos(self.frame_exclusivo_grafico)
                else:
                    self.grafico_barras(self.frame_exclusivo_grafico)
            elif tipo == 'pizza':
                self.checkbox_barras.deselect()
                self.checkbox_pizza.select()
                if self.checkbox_assunto.get():
                    self.atualizar_grafico_pizza_assuntos(self.frame_exclusivo_grafico)
                else:
                    self.atualizar_grafico_pizza(self.frame_exclusivo_grafico)
        except Exception as e:
            print(f"Erro ao mudar visualização: {e}")
            messagebox.showerror("Erro", f"Não foi possível mudar a visualização: {e}")
    
    def mudar_tema(self):
        tema = ctk.get_appearance_mode()

        if tema == "Dark":
            plt.style.use('dark_background')
            self.cor_fundo = '#2b2b2b'
            self.cor_texto = 'white'
            self.cor_bordas = 'white'
        else:
            plt.style.use('default')
            self.cor_fundo = 'white'
            self.cor_texto = 'black'
            self.cor_bordas = 'black'

    #Buscar os dados do banco de dados das áreas municipais!
    #Essa função, até a lina xxx, trata sobre a filtragem de quantidade POR ÁREAS MUNCIPAIS!
    #Para ver sobre a filtragem POR ASSUNTO (Problema, vá até a linha xxx)
    def buscar_dados_areas(self):
        """Busca dados das áreas municipais no banco de dados com filtro de data"""
        try:
            caminho_banco = Path("database/db/sistema_protocolos.db")
            conn = sqlite3.connect(caminho_banco)  
            cursor = conn.cursor()
    
            # Obter datas selecionadas no formato dd/mm/yyyy (mesmo formato armazenado)
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')
    
            # Verificar quais áreas estão selecionadas
            areas_selecionadas = [self.areas[i] for i, checkbox in enumerate(self.checkbox_areas) if checkbox.get()]
    
            if not areas_selecionadas:
                return {}  # Retorna dicionário vazio se nenhuma área estiver selecionada
    
            placeholders = ','.join(['?'] * len(areas_selecionadas))
            query = f"""
            SELECT area FROM protocolos 
            WHERE 
            substr(data_abertura, 7, 4) || substr(data_abertura, 4, 2) || substr(data_abertura, 1, 2) BETWEEN 
            substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2) AND 
            substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2)
            AND area IN ({placeholders})
            """
            params = [data_inicio, data_inicio, data_inicio, data_fim, data_fim, data_fim] + areas_selecionadas
    
            cursor.execute(query, params)
            resultados = cursor.fetchall()
    
            contador = Counter([area[0] for area in resultados])
            areas_ordenadas = dict(sorted(contador.items(), key=lambda item: item[1], reverse=True))
    
            return areas_ordenadas
    
        except Exception as e:
            print(f"Erro ao buscar dados: {e}")
            # Retorna dados fictícios apenas para as áreas selecionadas
            dados_ficticios = {
            'Centro-Sul': 1,
            'Sudoeste': 1,
            'Norte': 1,
            'Centro-Norte': 1,
            'Noroeste': 1,
            'Leste': 1
            }
            # Filtra apenas as áreas selecionadas
            areas_selecionadas = [self.areas[i] for i, checkbox in enumerate(self.checkbox_areas) if checkbox.get()]
            return {k: v for k, v in dados_ficticios.items() if k in areas_selecionadas}
        finally:
            if 'conn' in locals():
                conn.close()

    # Adicione este novo método à classe:
    def atualizar_grafico_por_checkbox(self, frame_grafico):
        """Atualiza o gráfico quando as checkboxes são alteradas"""
        # Verifica se há pelo menos uma área selecionada
        if not any(checkbox.get() for checkbox in self.checkbox_areas):
            messagebox.showwarning("Aviso", "Selecione pelo menos uma área municipal!")
            # Re-seleciona a última checkbox que foi desmarcada
            for checkbox in self.checkbox_areas:
                if not checkbox.get():
                    checkbox.select()
                break
            return
    
        # Verifica qual gráfico está ativo e atualiza
        if self.checkbox_barras.get():
            self.grafico_barras(frame_grafico)
        else:
            self.atualizar_grafico_pizza(frame_grafico)

    #Criar gráfico de baras
    def grafico_barras(self, frame_grafico):
        """Gera gráfico de barras com contagem de protocolos por área"""
        try:
            self.mudar_tema()  # Isso define self.cor_fundo, self.cor_texto, etc.
        
            # Remova plt.style.use('ggplot') para ter mais controle
            rcParams['font.family'] = 'Poppins'
            rcParams['font.size'] = 10

            # Obter dados
            dados = self.buscar_dados_areas()
            labels = list(dados.keys())
            valores = list(dados.values())

            # Criar figura
            fig = Figure(figsize=(8, 6), facecolor=self.cor_fundo)
            fig.subplots_adjust(bottom=0.25)
            ax = fig.add_subplot(111)
        
            # Configurar fundo do eixo e remover grades
            ax.set_facecolor(self.cor_fundo)
            ax.grid(False)  # Isso remove as grades

            # Configurar cores
            cmap = cm.get_cmap('tab10')
            cores = [cmap(i / len(labels)) for i in range(len(labels))]

            # Criar as barras com bordas
            barras = ax.bar(
            labels,
            valores,
            color=cores,
            edgecolor='black',  # Cor da borda
            linewidth=0.5,      # Espessura da borda
            width=0.7          # Largura das barras
            )

            # Configurar título
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')
            titulo = f'Protocolos por Área Municipal\nPeríodo: {data_inicio} a {data_fim}'
            ax.set_title(titulo, color=self.cor_texto, pad=20, fontsize=12)

            # Configurar eixos
            ax.set_ylabel('Quantidade', color=self.cor_texto)
            ax.tick_params(axis='x', colors=self.cor_texto, rotation=45, labelsize=10)
            ax.tick_params(axis='y', colors=self.cor_texto)

            # Configurar bordas do gráfico
            for spine in ['top', 'right']:
                ax.spines[spine].set_visible(False)
            for spine in ['bottom', 'left']:
                ax.spines[spine].set_color(self.cor_bordas)

            # Adicionar valores no topo das barras (em negrito)
            for barra, valor in zip(barras, valores):
                height = barra.get_height()
                ax.text(
                barra.get_x() + barra.get_width()/2.,
                height + 0.5,
                str(valor),
                ha='center',
                va='bottom',
                color=self.cor_texto,
                fontsize=11,
                fontweight='bold'  # Texto em negrito
                )

            # Limpar e atualizar o frame
            for widget in frame_grafico.winfo_children():
                widget.destroy()

            canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except Exception as e:
            print(f"Erro ao gerar gráfico de barras: {e}")
            # Mostrar mensagem de erro no frame
            for widget in frame_grafico.winfo_children():
                widget.destroy()
            
            label_erro = ctk.CTkLabel(
            frame_grafico,
            text="Erro ao gerar gráfico",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#FF5555"
            )
            label_erro.pack(expand=True)

    #Criar gráfico de PIZZA
    def atualizar_grafico_pizza(self, frame_grafico):
        """Atualiza o gráfico de pizza com as ocorrências por área municipal"""
        try:
            self.mudar_tema()

            dados = self.buscar_dados_areas()
            labels = list(dados.keys())
            valores = list(dados.values())
            
            rcParams['font.family'] = 'Poppins'

            def make_autopct(valores):
                def my_autopct(pct):
                    total = sum(valores)
                    valor_absoluto = int(round(pct * total / 100.0))
                    return f'({pct:.1f}%)'
                return my_autopct

            labels_formatadas = [f"{label} - {valores[i]}" for i, label in enumerate(labels)]

            fig = Figure(figsize=(6, 6), facecolor=self.cor_fundo)
            ax = fig.add_subplot(111)

            wedges, texts, autotexts = ax.pie(
                valores,
                labels=labels_formatadas,
                autopct=make_autopct(valores),
                startangle=90,
                wedgeprops={'linewidth': 0.5, 'edgecolor': 'black'},
                textprops=dict(color=self.cor_texto)
            )

            for autotext in autotexts:
                autotext.set_color('black')

            # Adiciona título com período selecionado
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')
            titulo = f'Distribuição por Área Municipal\nPeríodo: {data_inicio} a {data_fim}'
            
            ax.set_title(titulo, color=self.cor_texto, pad=20)

            fig.tight_layout()

            for widget in frame_grafico.winfo_children():
                widget.destroy()

            canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except Exception as e:
            print(f"Erro ao atualizar gráfico de pizza: {e}")
    
    
    #Calcular estatísticas gerais e ranking de bairros mais requisitados
    def calcular_estatisticas(self):
        """Calcula as estatísticas com base nos dados filtrados, incluindo ranking de bairros"""
        try:
            caminho_banco = Path("database/db/sistema_protocolos.db")
            conn = sqlite3.connect(caminho_banco)
            cursor = conn.cursor()
        
            # Obter datas selecionadas
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')
        
            # Consulta para obter o total de protocolos
            query_total = """
            SELECT COUNT(*) FROM protocolos 
            WHERE substr(data_abertura, 7, 4) || substr(data_abertura, 4, 2) || substr(data_abertura, 1, 2) BETWEEN 
              substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2) AND 
              substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2)
            """
            cursor.execute(query_total, [data_inicio]*3 + [data_fim]*3)
            total = cursor.fetchone()[0]
        
            # Consulta para obter a área mais requisitada
            query_area = """
            SELECT area, COUNT(*) as count FROM protocolos 
            WHERE substr(data_abertura, 7, 4) || substr(data_abertura, 4, 2) || substr(data_abertura, 1, 2) BETWEEN 
              substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2) AND 
              substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2)
            GROUP BY area ORDER BY count DESC LIMIT 1
            """
            cursor.execute(query_area, [data_inicio]*3 + [data_fim]*3)
            area_mais = cursor.fetchone()
            area_mais = area_mais[0] if area_mais else "Nenhum dado"
        
            # Consulta para obter ranking de bairros (top 5)
            query_bairros = """
            SELECT bairro, COUNT(*) as count FROM protocolos 
            WHERE bairro IS NOT NULL AND bairro != ''
              AND substr(data_abertura, 7, 4) || substr(data_abertura, 4, 2) || substr(data_abertura, 1, 2) BETWEEN 
                substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2) AND 
                substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2)
            GROUP BY bairro ORDER BY count DESC LIMIT 5
            """
            cursor.execute(query_bairros, [data_inicio]*3 + [data_fim]*3)
            ranking_bairros = cursor.fetchall()
        
            # Calcular diferença de dias entre as datas
            delta = self.data_ate.get_date() - self.data_de.get_date()
            dias = delta.days + 1  # +1 para incluir o dia inicial
            semanas = max(dias / 7, 1)  # Pelo menos 1 semana
        
            # Calcular médias
            media_diaria = round(total / dias, 2) if dias > 0 else 0
            media_semanal = round(total / semanas, 2)
        
            return {
                'total': total,
                'area_mais': area_mais,
                'media_diaria': media_diaria,
                'media_semanal': media_semanal,
                'ranking_bairros': ranking_bairros  # Lista de tuplas (bairro, quantidade)
            }
        
        except Exception as e:
            print(f"Erro ao calcular estatísticas: {e}")
        finally:
            if 'conn' in locals():
                conn.close()

    def atualizar_resultados(self, frame):
        """Atualiza os labels com as estatísticas calculadas e o ranking de bairros"""
        estatisticas = self.calcular_estatisticas()

        # Atualizar estatísticas gerais
        self.label_total.configure(text=f"Total do período: {estatisticas['total']}")
        self.label_area_mais.configure(text=f"Área mais requisitada: {estatisticas['area_mais']}")
        self.label_media_semanal.configure(text=f"Média semanal: {estatisticas['media_semanal']}")
        self.label_media_diaria.configure(text=f"Média diária: {estatisticas['media_diaria']}")

        # Atualizar ranking de bairros (top 5)
        for i in range(5):  # Alterado para 5, que é o número de itens no ranking
            if i < len(estatisticas['ranking_bairros']):
                bairro, qtd = estatisticas['ranking_bairros'][i]
                posicao, conteudo = self.labels_ranking[i]
                conteudo.configure(text=f"{bairro}: {qtd}")
            else:
                posicao, conteudo = self.labels_ranking[i]
                conteudo.configure(text="")
    
    #Filtragem de dados por assunto

    def alternar_modo_visualizacao(self):
        """Alterna entre visualização por área e por assunto"""
        try:
            if self.checkbox_assunto.get():
                # Desmarca o checkbox "Por quantidade"
                self.checkbox_quantidade.deselect()
            
                # Esconder o container de áreas
                self.container_areas.pack_forget()

                # Atualizar o gráfico para mostrar por assunto
                if self.checkbox_barras.get():
                    self.grafico_barras_assuntos(self.frame_exclusivo_grafico)
                else:
                    self.atualizar_grafico_pizza_assuntos(self.frame_exclusivo_grafico)
            else:
                # Marca o checkbox "Por quantidade"
                self.checkbox_quantidade.select()
                
                # Mostrar o container de áreas
                self.container_areas.pack(side="left", fill="x", expand=True)

                # Atualizar o gráfico para mostrar por área
                if self.checkbox_barras.get():
                    self.grafico_barras(self.frame_exclusivo_grafico)
                else:
                    self.atualizar_grafico_pizza(self.frame_exclusivo_grafico)
                
        except Exception as e:
            print(f"Erro ao alternar modo de visualização: {e}")

    def buscar_dados_assuntos(self):
        """Busca dados de assuntos no banco de dados com filtro de data"""
        try:
            caminho_banco = Path("database/db/sistema_protocolos.db")
            conn = sqlite3.connect(caminho_banco)
            cursor = conn.cursor()

            # Obter datas selecionadas no formato dd/mm/yyyy
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')

            query = """
            SELECT problema, COUNT(*) as count FROM protocolos 
              WHERE substr(data_abertura, 7, 4) || substr(data_abertura, 4, 2) || substr(data_abertura, 1, 2) BETWEEN 
              substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2) AND 
              substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2)
              AND problema IS NOT NULL  
            GROUP BY problema ORDER BY count DESC
            """
            cursor.execute(query, [data_inicio]*3 + [data_fim]*3)
            resultados = cursor.fetchall()

            # Filtra novamente para garantir que não há problemas None
            assuntos_ordenados = {
            str(item[0]): item[1] for item in resultados 
            if item[0] is not None  # Garante que o problema não é None
            }

            return assuntos_ordenados

        except Exception as e:
            print(f"Erro ao buscar dados de assuntos: {e}")
            
        finally:
            if 'conn' in locals():
                conn.close()

    def grafico_barras_assuntos(self, frame_grafico):
        """Gera gráfico de barras com contagem de protocolos por assunto"""
        try:
            # Configuração do tema e estilo
            plt.style.use('default')  # Removemos o estilo ggplot
            rcParams['font.family'] = 'Poppins'
            rcParams['font.size'] = 9
            
            # Obter e preparar os dados
            dados = self.buscar_dados_assuntos()
            if not dados:
                raise ValueError("Nenhum dado encontrado")
                
            # Processar labels longos (agora com reticências)
            def formatar_label(label):
                max_len = 24  # Número máximo de caracteres
                return label[:max_len] + '...' if len(label) > max_len else label
            
            labels = [formatar_label(k) for k in dados.keys()]
            valores = list(dados.values())

            # Configurar cores
            cmap = cm.get_cmap('tab10')
            cores = [cmap(i / len(labels)) for i in range(len(labels))]

            # Criar figura com fundo adaptável
            fig = Figure(figsize=(8, 6), facecolor=self.cor_fundo)
            fig.subplots_adjust(bottom=0.35, top=0.85)
            ax = fig.add_subplot(111)
            
            # Configurar fundo do gráfico
            ax.set_facecolor(self.cor_fundo)
            
            # Criar as barras
            barras = ax.bar(labels, valores, color=cores, width=0.6, 
                          edgecolor=self.cor_bordas, linewidth=0.5)

            # Configurar título
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')
            titulo = f'Protocolos por Tipo de Assunto\n{data_inicio} a {data_fim}'
            ax.set_title(titulo, color=self.cor_texto, pad=20, fontsize=12)

            # Configurar eixos
            ax.set_ylabel('Quantidade', color=self.cor_texto)
            ax.tick_params(axis='y', colors=self.cor_texto)
            ax.tick_params(axis='x', colors=self.cor_texto)
            
            # Configurar labels do eixo X
            ax.set_xticks(range(len(labels)))
            ax.set_xticklabels(
                labels, 
                rotation=45, 
                ha='right', 
                rotation_mode='anchor',
                fontsize=9
            )
            
            # Remover grade e ajustar bordas
            ax.grid(False)  # Remove a grade de fundo
            for spine in ax.spines.values():
                spine.set_color(self.cor_bordas)

            # Adicionar valores no topo das barras
            for barra, valor in zip(barras, valores):
                height = barra.get_height()
                ax.text(
                    barra.get_x() + barra.get_width()/2., 
                    height + 0.5, 
                    str(valor), 
                    ha='center', 
                    va='bottom', 
                    color=self.cor_texto,
                    fontsize=9,
                    fontweight='bold'
                )

            # Limpar e atualizar o frame
            for widget in frame_grafico.winfo_children():
                widget.destroy()

            canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except ValueError as e:
            for widget in frame_grafico.winfo_children():
                widget.destroy()
                
            label_erro = ctk.CTkLabel(
                frame_grafico,
                text="Nenhum dado encontrado\npara os filtros selecionados",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#FF5555"
            )
            label_erro.pack(expand=True)

        except Exception as e:
            print(f"Erro ao gerar gráfico: {e}")
            for widget in frame_grafico.winfo_children():
                widget.destroy()
                
            label_erro = ctk.CTkLabel(
                frame_grafico,
                text="Erro ao gerar gráfico",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#FF5555"
            )
            label_erro.pack(expand=True)

    def atualizar_grafico_pizza_assuntos(self, frame_grafico):
        """Atualiza o gráfico de pizza com as ocorrências por assunto"""
        try:
            self.mudar_tema()
            plt.style.use('ggplot')
            rcParams['font.family'] = 'Poppins'
            rcParams['font.size'] = 9  # Tamanho da fonte base

            dados = self.buscar_dados_assuntos()
            if not dados:
                raise ValueError("Nenhum dado encontrado para os filtros selecionados")

            # Filtra e prepara os dados
            dados = {k: v for k, v in dados.items() if v > 0}
            labels = list(dados.keys())
            valores = list(dados.values())

            if not valores:
                raise ValueError("Dados insuficientes para gerar o gráfico")

            # Agrupa itens menores em "Outros" se houver muitos
            if len(labels) > 8:
                outros = sum(valores[8:])
                labels = labels[:8] + ['Outros']
                valores = valores[:8] + [outros]

            # Aumentamos o tamanho da figura aqui (de 6,5 para 7,5)
            fig = Figure(figsize=(7.5, 6), facecolor=self.cor_fundo)  # Tamanho maior
            fig.subplots_adjust(left=0.1, right=0.7)  # Ajuste de margens
            
            ax = fig.add_subplot(111)

            # Função para mostrar valores absolutos e percentuais
            def make_autopct(valores):
                def my_autopct(pct):
                    total = sum(valores)
                    valor = int(round(pct * total / 100.0))
                    return f'{valor} ({pct:.1f}%)' if pct > 5 else ''
                return my_autopct

            # Cria o gráfico de pizza
            wedges, texts, autotexts = ax.pie(
                valores,
                labels=labels,
                autopct=make_autopct(valores),
                startangle=90,
                pctdistance=0.8,
                labeldistance=1.05,
                textprops={'color': self.cor_texto, 'fontsize': 10},  # Fonte um pouco maior
                wedgeprops={'linewidth': 0.5, 'edgecolor': self.cor_bordas}
            )

            # Configura o título
            data_inicio = self.data_de.get_date().strftime('%d/%m/%Y')
            data_fim = self.data_ate.get_date().strftime('%d/%m/%Y')
            titulo = f'Distribuição por Assunto\n{data_inicio} a {data_fim}'
            ax.set_title(titulo, color=self.cor_texto, pad=20, fontsize=12)

            # Limpa o frame e adiciona o novo gráfico
            for widget in frame_grafico.winfo_children():
                widget.destroy()

            canvas = FigureCanvasTkAgg(fig, master=frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)

        except ValueError as e:
            # Tratamento de erro (mantido igual)
            for widget in frame_grafico.winfo_children():
                widget.destroy()
            
            label_erro = ctk.CTkLabel(
                frame_grafico,
                text=f"Nenhum dado encontrado\npara os filtros selecionados",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#FF5555"
            )
            label_erro.pack(expand=True)

        except Exception as e:
            # Tratamento de erro genérico (mantido igual)
            for widget in frame_grafico.winfo_children():
                widget.destroy()
                
            label_erro = ctk.CTkLabel(
                frame_grafico,
                text="Erro ao gerar gráfico",
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color="#FF5555"
            )
            label_erro.pack(expand=True)

    #Janela com o ranking completo dos bairos
    def mostrar_todos_bairros(self):
        """Exibe todos os bairros em uma janela MDI com tabela, incluindo a área municipal"""
        try:
            # Obter dados filtrados
            ranking_completo = self.buscar_todos_bairros(
                self.data_de.get_date().strftime('%d/%m/%Y'),
                self.data_ate.get_date().strftime('%d/%m/%Y')
            )
            
            # Criar janela MDI
            janela_mdi = ctk.CTkToplevel(self)
            janela_mdi.title("Ranking Completo de Bairros")
            janela_mdi.transient(self)
            janela_mdi.grab_set()
            janela_mdi.resizable(False, False)
            janela_mdi.attributes('-toolwindow', True)
            
            # Configurar tamanho e posição
            largura = 800  # Aumentei a largura para acomodar a nova coluna
            altura = 500
            x = self.winfo_x() + (self.winfo_width() - largura) // 2
            y = self.winfo_y() + (self.winfo_height() - altura) // 2
            janela_mdi.geometry(f"{largura}x{altura}+{x}+{y}")
            
            # Frame principal
            frame_principal = ctk.CTkFrame(janela_mdi, corner_radius=10)
            frame_principal.pack(padx=10, pady=10, fill="both", expand=True)
            
            # Título
            label_titulo = ctk.CTkLabel(
                frame_principal,
                text="Ranking Completo de Bairros",
                font=ctk.CTkFont("Poppins", size=14, weight="bold")
            )
            label_titulo.pack(pady=(10, 5))
            
            # Período
            label_periodo = ctk.CTkLabel(
                frame_principal,
                text=f"Período: {self.data_de.get_date().strftime('%d/%m/%Y')} a {self.data_ate.get_date().strftime('%d/%m/%Y')}",
                font=ctk.CTkFont("Poppins", size=12)
            )
            label_periodo.pack(pady=(0, 10))
            
            # Barra de pesquisa
            frame_pesquisa = ctk.CTkFrame(frame_principal, fg_color="transparent")
            frame_pesquisa.pack(fill="x", padx=10, pady=(0, 10))
            
            entrada_pesquisa = ctk.CTkEntry(
                frame_pesquisa,
                placeholder_text="Pesquise por bairro ou área...",
                font=ctk.CTkFont("Poppins", 12)
            )
            entrada_pesquisa.pack(side="left", fill="x", expand=True)
            
            # Criar Treeview (tabela) com estilo adaptável
            style = ttk.Style()
            style.theme_use("clam")
            
            def atualizar_estilo():
                modo_claro = ctk.get_appearance_mode() == "Light"
                
                cor_fundo = "#f5f5f5" if modo_claro else "#2b2b2b"
                cor_texto = "#000000" if modo_claro else "#ffffff"
                cor_cabecalho = "#e1e1e1" if modo_claro else "#3a3a3a"
                cor_selecao = "#0078D7"
                
                style.configure("Treeview",
                              background=cor_fundo,
                              foreground=cor_texto,
                              fieldbackground=cor_fundo,
                              bordercolor="#d3d3d3" if modo_claro else "#3a3a3a",
                              borderwidth=0)
                style.configure("Treeview.Heading",
                              background=cor_cabecalho,
                              foreground=cor_texto,
                              font=('Poppins', 10, 'bold'))
                style.map("Treeview",
                          background=[('selected', cor_selecao)],
                          foreground=[('selected', 'white')])
                
                if modo_claro:
                    tabela.tag_configure('oddrow', background='white')
                    tabela.tag_configure('evenrow', background='#f0f0f0')
                else:
                    tabela.tag_configure('oddrow', background='#2b2b2b')
                    tabela.tag_configure('evenrow', background='#333333')
            
            # Frame para a tabela com scrollbar
            frame_tabela = ctk.CTkFrame(frame_principal, fg_color="transparent")
            frame_tabela.pack(fill="both", expand=True, padx=10, pady=(0, 10))
            
            # Scrollbar vertical
            scroll_y = ttk.Scrollbar(frame_tabela)
            scroll_y.pack(side="right", fill="y")
            
            # Treeview (tabela) com coluna adicional para a área
            tabela = ttk.Treeview(
                frame_tabela,
                columns=("posicao", "bairro", "area", "quantidade"),
                show="headings",
                yscrollcommand=scroll_y.set,
                selectmode="browse",
                height=15
            )
            tabela.pack(fill="both", expand=True)
            
            # Configurar scrollbar
            scroll_y.config(command=tabela.yview)
            
            # Configurar colunas
            tabela.heading("posicao", text="Posição", anchor="center")
            tabela.heading("bairro", text="Bairro", anchor="w")
            tabela.heading("area", text="Área Municipal", anchor="w")
            tabela.heading("quantidade", text="Protocolos", anchor="center")
            
            tabela.column("posicao", width=60, anchor="center")
            tabela.column("bairro", width=250, anchor="w")
            tabela.column("area", width=200, anchor="w")
            tabela.column("quantidade", width=80, anchor="center")
            
            # Adicionar dados à tabela
            dados_originais = []
            for i, (bairro, area, qtd) in enumerate(ranking_completo, start=1):
                dados_originais.append((f"{i}º", bairro, area, qtd))
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                tabela.insert("", "end", values=(f"{i}º", bairro, area, qtd), tags=(tag,))
            
            # Função de filtro
            def filtrar_bairros(event=None):
                termo = entrada_pesquisa.get().lower()
                tabela.delete(*tabela.get_children())
                
                for i, (posicao, bairro, area, qtd) in enumerate(dados_originais, start=1):
                    if (termo in bairro.lower()) or (termo in area.lower()):
                        tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                        tabela.insert("", "end", values=(posicao, bairro, area, qtd), tags=(tag,))
            
            entrada_pesquisa.bind("<KeyRelease>", filtrar_bairros)
            
            # Configurar estilo inicial
            atualizar_estilo()
            
            # Observar mudanças no tema
            janela_mdi.bind("<Configure>", lambda e: atualizar_estilo())
            ctk.AppearanceModeTracker.add(atualizar_estilo)
            
        except Exception as e:
            print(f"Erro ao mostrar ranking completo: {e}")
            messagebox.showerror("Erro", "Não foi possível carregar o ranking completo")

    # Método para buscar todos os bairros (sem limite)
    def buscar_todos_bairros(self, data_inicio=None, data_fim=None):
        """Busca todos os bairros no banco de dados com filtro de datas, incluindo a área municipal"""
        try:
            caminho_banco = Path("database/db/sistema_protocolos.db")
            conn = sqlite3.connect(caminho_banco)
            cursor = conn.cursor()
        
            query = """
            SELECT bairro, area, COUNT(*) as count FROM protocolos 
            WHERE bairro IS NOT NULL AND bairro != ''
            AND substr(data_abertura, 7, 4) || substr(data_abertura, 4, 2) || substr(data_abertura, 1, 2) BETWEEN 
            substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2) AND 
            substr(?, 7, 4) || substr(?, 4, 2) || substr(?, 1, 2)
            GROUP BY bairro, area ORDER BY count DESC
            """
            cursor.execute(query, [data_inicio]*3 + [data_fim]*3)
            resultados = cursor.fetchall()
        
            # Retorna uma lista de tuplas (bairro, área, quantidade)
            return resultados
            
        except Exception as e:
            print(f"Erro ao buscar todos os bairros: {e}")
            return []
        finally:
            if 'conn' in locals():
                conn.close()
    
    def alternar_tema(self):
        """Alterna entre os temas claro e escuro e atualiza os gráficos"""
        try:
            # Alternar o tema global
            tema_atual = ctk.get_appearance_mode()
            novo_tema = "Light" if tema_atual == "Dark" else "Dark"
            ctk.set_appearance_mode(novo_tema)
        
            # Atualizar o ícone do botão
            novo_icone = self.icone_sol if novo_tema == "Dark" else self.icone_lua
            self.botao_tema.configure(image=novo_icone)
            
        
            # Atualizar cores dos gráficos
            self.mudar_tema()
        
            # Forçar a atualização dos gráficos
            if self.checkbox_barras.get():
                if self.checkbox_assunto.get():
                    self.grafico_barras_assuntos(self.frame_exclusivo_grafico)
                else:
                    self.grafico_barras(self.frame_exclusivo_grafico)
            else:
                if self.checkbox_assunto.get():
                    self.atualizar_grafico_pizza_assuntos(self.frame_exclusivo_grafico)
                else:
                    self.atualizar_grafico_pizza(self.frame_exclusivo_grafico)
                
        except Exception as e:
            print(f"Erro ao alternar tema: {e}")
    
    #fechar suavemente a janela de analise de dados
    def fechar_janela_analise(self):
        def fade_out(opacity):
            if opacity > 0:
                self.attributes("-alpha", opacity)
                self.after(9, lambda: fade_out(opacity - 0.05))
            else:
                self.destroy()

        fade_out(1.0)

