# Este sistema foi criado por Jadson Pamplona Viana

import pandas as pd
import matplotlib.pyplot as plt
from tkinter import Tk, filedialog, messagebox, Label, Button, Listbox, Scrollbar, SINGLE
from tkcalendar import DateEntry
import tkinter as tk
from datetime import datetime
import locale
import tkinter as tk
from PIL import Image, ImageTk

# Configurar local para português do Brasil
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


class PlanilhaWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistema de Análise de Big Data")
        self.root.geometry("400x200")
        self.root.state('zoomed')  # Maximiza a janela principal

        try:
            # Carregar a imagem com Pillow
            original_image = Image.open("polo.png")

            # Converter a imagem para um formato adequado para Tkinter
            background_image = ImageTk.PhotoImage(original_image)

            # Criar um label para a imagem de fundo e configurá-lo para preencher a janela
            self.background_label = tk.Label(root, image=background_image)
            # Preencher toda a janela com a imagem de fundo
            self.background_label.place(x=0, y=0, relwidth=1, relheight=1)

            # Garantir que a imagem de fundo permaneça atrás dos outros widgets
            self.background_label.lower()

            # Manter uma referência à imagem para evitar que seja coletada pelo garbage collector
            self.background_label.image = background_image

        except Exception as e:
            print(f"Erro ao carregar a imagem: {e}")

        self.label = Label(
            root, text="Selecione a análise desejada:", font=("Arial", 16, "bold"))
        self.label.pack(pady=20)

        self.button1 = Button(root, text="Total Devido por Obra", font=("Arial", 12), bg='#007ACC', fg='white',
                              activebackground='#005D99', padx=20, pady=10,
                              command=self.selecionar_planilha_obras)
        self.button1.pack(pady=10)

        self.button2 = Button(root, text="Valor Devido por Obra + Data", font=("Arial", 12), bg='#2E8B57', fg='white',
                              activebackground='#226E42', padx=20, pady=10,
                              command=self.abrir_tela_valor_por_data)
        self.button2.pack(pady=10)

        self.button3 = Button(root, text="Total Devido por Parcelas + Obra", font=("Arial", 12), bg='#8B0000', fg='white',
                              activebackground='#6E0000', padx=20, pady=10,
                              command=self.selecionar_planilha_parcelas)
        self.button3.pack(pady=10)

    def selecionar_planilha_obras(self):
        self.root.withdraw()  # Oculta a janela atual
        try:
            filepath = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if not filepath:
                self.root.deiconify()  # Mostra a janela anterior se o usuário cancelar a seleção
                # Garantir que a janela fique maximizada
                self.root.state('zoomed')
                return

            df = pd.read_excel(filepath)
            agrupar_dados_e_gerar_grafico(
                df, 'Obra', 'Total Devido por Obra', self.root)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
            self.root.deiconify()  # Mostra a janela anterior em caso de erro
            self.root.state('zoomed')  # Garantir que a janela fique maximizada

    def abrir_tela_valor_por_data(self):
        self.root.withdraw()  # Oculta a janela principal
        window = TelaValorPorData(self.root)

    def selecionar_planilha_parcelas(self):
        self.root.withdraw()  # Oculta a janela atual
        try:
            filepath = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if not filepath:
                self.root.deiconify()  # Mostra a janela anterior se o usuário cancelar a seleção
                # Garantir que a janela fique maximizada
                self.root.state('zoomed')
                return

            df = pd.read_excel(filepath)
            self.abrir_tela_selecao_obra(df)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
            self.root.deiconify()  # Mostra a janela anterior em caso de erro
            self.root.state('zoomed')  # Garantir que a janela fique maximizada

    def abrir_tela_selecao_obra(self, df):
        self.root.withdraw()  # Oculta a janela principal
        window = TelaSelecaoObra(self.root, df)


class TelaValorPorData:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()  # Oculta a janela principal
        self.toplevel = tk.Toplevel(root)
        self.toplevel.title("Valor Devido por Obra + Data")
        self.toplevel.geometry("300x250")
        self.toplevel.state('zoomed')  # Maximiza a janela secundária

        self.label_inicio = Label(
            self.toplevel, text="Data de Início:", font=("Arial", 14, "bold"))
        self.label_inicio.pack(pady=10)

        self.entry_inicio = DateEntry(self.toplevel, font=(
            "Arial", 12), date_pattern='dd/MM/yyyy')
        self.entry_inicio.pack(pady=5)

        self.label_fim = Label(
            self.toplevel, text="Data de Fim:", font=("Arial", 14, "bold"))
        self.label_fim.pack(pady=10)

        self.entry_fim = DateEntry(self.toplevel, font=(
            "Arial", 12), date_pattern='dd/MM/yyyy')
        self.entry_fim.pack(pady=5)

        self.button_confirmar = Button(self.toplevel, text="Confirmar", font=("Arial", 12), bg='#2E8B57', fg='white',
                                       activebackground='#226E42', padx=20, pady=10,
                                       command=self.validar_datas_fechar_janela)
        self.button_confirmar.pack(pady=20)

    def validar_datas_fechar_janela(self):
        start_date = self.entry_inicio.get()
        end_date = self.entry_fim.get()

        try:
            datetime.strptime(start_date, "%d/%m/%Y")
            datetime.strptime(end_date, "%d/%m/%Y")

            filepath = filedialog.askopenfilename(
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
            if not filepath:
                self.root.deiconify()  # Mostra a janela principal se o usuário cancelar a seleção
                # Garantir que a janela fique maximizada
                self.root.state('zoomed')
                return

            df = pd.read_excel(filepath)
            agrupar_por_obra_e_data(df, start_date, end_date, self.root)
            self.toplevel.destroy()  # Fecha a janela após processar os dados

        except ValueError:
            messagebox.showerror(
                "Erro", "Formato de data inválido. Utilize o formato DD/MM/AAAA.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao validar as datas: {e}")


class TelaSelecaoObra:
    def __init__(self, root, df):
        self.root = root
        self.df = df
        self.root.withdraw()  # Oculta a janela principal
        self.toplevel = tk.Toplevel(root)
        self.toplevel.title("Seleção de Obra")
        self.toplevel.geometry("400x300")
        self.toplevel.state('zoomed')  # Maximiza a janela secundária

        self.label = Label(
            self.toplevel, text="Selecione a Obra:", font=("Arial", 16, "bold"))
        self.label.pack(pady=20)

        self.scrollbar = Scrollbar(self.toplevel)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = Listbox(self.toplevel, font=(
            "Arial", 14), yscrollcommand=self.scrollbar.set, selectmode=SINGLE)
        self.listbox.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.listbox.yview)

        obras = df['Obra'].unique()
        for obra in obras:
            self.listbox.insert(tk.END, obra)

        self.button_confirmar = Button(self.toplevel, text="Confirmar", font=("Arial", 12), bg='#8B0000', fg='white',
                                       activebackground='#6E0000', padx=20, pady=10,
                                       command=self.selecionar_obra)
        self.button_confirmar.pack(pady=20)

    def selecionar_obra(self):
        selected_index = self.listbox.curselection()
        if not selected_index:
            messagebox.showerror("Erro", "Nenhuma obra selecionada.")
            return

        selected_obra = self.listbox.get(selected_index)
        self.df_filtrada = self.df[self.df['Obra'] == selected_obra]
        agrupar_por_parcelas_e_obra(self.df_filtrada, selected_obra, self.root)
        self.toplevel.destroy()  # Fecha a janela após processar os dados


def agrupar_dados_e_gerar_grafico(df, groupby_column, title, root):
    try:
        # Agrupando os dados pela coluna especificada e somando os valores da coluna 'Vlr. Parcela'
        grouped_df = df.groupby(groupby_column)[
            'Vlr. Parcela'].sum().sort_values()

        total_value = grouped_df.sum()

        # Definindo cores para o gráfico
        colors = ['#4C72B0', '#55A868', '#C44E52', '#8172B2', '#CCB974']
        bar_colors = [colors[i % len(colors)] for i in range(len(grouped_df))]

        # Gerando o gráfico de barras horizontais com borda mais larga e cores variadas
        fig, ax = plt.subplots(figsize=(10, 6))
        bars = ax.barh(grouped_df.index.astype(
            str), grouped_df.values, color=bar_colors, linewidth=1.5)

        ax.set_title(title, fontsize=20, pad=20, color='#333333')
        ax.set_ylabel(groupby_column, fontsize=16, color='#333333')

        # Adicionando rótulos de valor e porcentagem
        for bar in bars:
            width = bar.get_width()
            percentage = (width / total_value) * 100
            label = f'{locale.currency(width, grouping=True)} ({
                percentage:.2f}%)'
            if width < ax.get_xlim()[1] / 10:
                ax.text(width * 1.02, bar.get_y() + bar.get_height() / 2,
                        label, ha='left', va='center', color='black', fontsize=12)
            else:
                ax.text(width / 2, bar.get_y() + bar.get_height() / 2, label,
                        ha='center', va='center', color='black', fontsize=12)

        # Adicionando o valor total fora das barras
        total_label = f'Total Devido {
            locale.currency(total_value, grouping=True)}'
        plt.figtext(0.5, 0.01, total_label, ha='center',
                    fontsize=14, weight='bold', color='#333333')

        # Maximizando a janela do gráfico
        manager = plt.get_current_fig_manager()
        manager.window.state('zoomed')

        plt.tight_layout()
        plt.show()
        root.deiconify()  # Mostra a janela anterior após fechar o gráfico
        root.state('zoomed')  # Garantir que a janela fique maximizada

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar os dados: {e}")
        root.deiconify()  # Mostra a janela anterior em caso de erro
        root.state('zoomed')  # Garantir que a janela fique maximizada


def agrupar_por_obra_e_data(df, start_date, end_date, root):
    try:
        # Convertendo as strings de data para objetos datetime
        start_date = datetime.strptime(start_date, "%d/%m/%Y")
        end_date = datetime.strptime(end_date, "%d/%m/%Y")

        # Filtrando os dados dentro do intervalo de datas
        df['Dt. Vencimento'] = pd.to_datetime(
            df['Dt. Vencimento'], errors='coerce')
        filtered_df = df[(df['Dt. Vencimento'] >= start_date)
                         & (df['Dt. Vencimento'] <= end_date)]

        agrupar_dados_e_gerar_grafico(filtered_df, 'Obra',
                                      f'Valor por Obra entre {start_date.strftime(
                                          "%d/%m/%Y")} e {end_date.strftime("%d/%m/%Y")}',
                                      root)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar os dados: {e}")
        root.deiconify()  # Mostra a janela anterior em caso de erro
        root.state('zoomed')  # Garantir que a janela fique maximizada


def agrupar_por_parcelas_e_obra(df, obra, root):
    try:
        # Agrupando os dados pelo número de parcelas atrasadas e somando os valores da coluna 'Valor Devido'
        grouped_df = df.groupby('Nº Parc. Atraso').agg(
            {'Valor Devido': 'sum', 'Venda': 'count'}).sort_index()

        total_value = grouped_df['Valor Devido'].sum()
        total_vendas = grouped_df['Venda'].sum()

        # Definindo cores e estilos corporativos para o gráfico
        colors = ['#4C72B0', '#55A868', '#C44E52', '#8172B2', '#CCB974']
        bar_colors = [colors[i % len(colors)] for i in range(len(grouped_df))]

        # Gerando o gráfico de barras horizontais com borda mais larga e cores personalizadas
        # Aumentando o tamanho do gráfico para mais espaço
        fig, ax = plt.subplots(figsize=(14, 8))
        bars = ax.barh(grouped_df.index.astype(
            str), grouped_df['Valor Devido'], color=bar_colors, linewidth=1.5)

        ax.set_title(f'Total Devido por Nº de Parcelas Atrasadas para a Obra {obra}', fontsize=20, pad=20, color='#333333')
        ax.set_ylabel('Nº Parc. Atraso', fontsize=16, color='#333333')

        # Adicionando margem extra no eixo X para evitar invasão dos rótulos
        max_valor_devido = grouped_df['Valor Devido'].max()
        # Adicionando 30% de espaço adicional
        ax.set_xlim(0, max_valor_devido * 1.3)

        # Adicionando rótulos de valor, porcentagem e quantidade de vendas
        for bar, venda in zip(bars, grouped_df['Venda']):
            width = bar.get_width()
            percentage = (width / total_value) * 100
            label = f'{locale.currency(width, grouping=True)} ({percentage:.2f}%) - {venda} vendas'

            # Sempre posicionar os rótulos fora da barra para evitar invasão
            ax.text(width * 1.02, bar.get_y() + bar.get_height() / 2,
                    label, ha='left', va='center', color='black', fontsize=12)

        # Adicionando o valor total fora das barras, com um pequeno espaço extra
        total_label = f'Total Devido: {locale.currency(total_value, grouping=True)} - Total de Vendas: {total_vendas}'

        # Usando plt.annotate para adicionar o rótulo total abaixo da barra mais à direita
        ax.annotate(total_label,
                    # posição centralizada abaixo da área de plotagem
                    xy=(0.5, -0.12),
                    xycoords='axes fraction',  # coordenadas fracionais
                    fontsize=14, weight='bold', color='#333333',
                    ha='center', va='center')

        # Maximizando a janela do gráfico
        manager = plt.get_current_fig_manager()
        manager.window.state('zoomed')

        plt.tight_layout(pad=3)  # Adicionando um pequeno padding extra
        plt.show()
        root.deiconify()  # Mostra a janela anterior após fechar o gráfico
        root.state('zoomed')  # Garantir que a janela fique maximizada

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar os dados: {e}")
        root.deiconify()  # Mostra a janela anterior em caso de erro
        root.state('zoomed')  # Garantir que a janela fique maximizada



if __name__ == "__main__":
    root = tk.Tk()
    planilha_window = PlanilhaWindow(root)
    root.mainloop()
