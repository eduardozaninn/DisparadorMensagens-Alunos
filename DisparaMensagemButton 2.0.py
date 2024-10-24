





# Importar bibliotecas necessárias
import tkinter as tk
from tkinter import ttk
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style    
from urllib.parse import quote
from time import sleep
from datetime import datetime
import pandas as pd
import webbrowser
import threading
import pyautogui
import json
import os
import logging





#================================================================================================================================



logging.basicConfig(filename='mesagens_enviadas.log', level=logging.INFO, format='%(asctime)s - %(message)s')



#================================================================================================================================



# Função para carregar o histórico de linha da última mensagem enviada
def load_last_line():
    if os.path.exists("last_line.json"):
        with open("last_line.json", "r") as f:
            data = json.load(f)
            return data.get("Ultima_linha_enviada", 0) # Retorna 0 se não encontrar a chave "ultima_linha"
    return 0

# Função para salvar a última mensagem enviada no histórico
def save_last_line(last_line):
    with open("last_line.json", "w") as f:
        json.dump({"Ultima_linha_enviada": last_line}, f)



# carregar as configurações do arquivo JSON
def load_settings():
    if os.path.exists("settings.json"):
        with open("settings.json", "r") as f:
            return json.load(f)
    else:
        return {"theme": "journal"} 
    


# salvar as configurações no arquivo JSON
def save_settings(settings):
    with open("settings.json", "w") as f:
        json.dump(settings, f)


# ================================================================================================================================



# Criar uma classe para a GUI
class CourseOfferGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Disparador de Mensagem Cursos")
        self.root.geometry("590x590")

        

        # Carregar configurações
        self.settings = load_settings()

        # Aplicar o tema carregado
        self.style = Style(theme=self.settings["theme"])

        # Carregar a última linha enviada
        self.last_line = load_last_line()



        # ================================================================================================================================



        # Lista os temas disponíveis
        available_themes = ttk.Style().theme_names()
        print("Temas disponíveis:", available_themes)

        # ADICIONA UMA OPÇÃO DE MENU PARA ESCOLHER O TEMA
        self.menu_theme = ttk.Menubutton(root, text="Escolher Tema")
        self.theme = tk.Menu(self.menu_theme, tearoff=-1)

        themes = ["superhero", "flatly", "darkly", "journal", "cyborg", "lumen", "minty", "pulse", "sandstone", "solar", "united", "yeti", "cerulean", "cosmo", "litera", "morph", "simplex", "vapor"]

        for theme in themes:
            self.theme.add_command(label=theme.capitalize(), command=lambda t=theme: self.change_theme(t))

        self.menu_theme.configure(menu=self.theme)
        self.menu_theme.pack(pady=0)



        # ================================================================================================================================

        # Criação de um frame principal para organizar os widgets
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=10, pady=10)

        #  título da aplicação
        self.title_label = ttk.Label(self.root, text='Disparador de Mensagem Cursos', font=('Arial', 20))
        self.title_label.pack(pady=25)

        # ================================================================================================================================



        # Recebe os dados de entrada.
        self.menu_course = ttk.Menubutton(root, text="Curso")
        self.course_selected = tk.StringVar()
        self.course = tk.Menu(self.menu_course, tearoff=0)



        # Adicione os cursos aqui
        self.course.add_command(label="INTRODUÇÃO A INFORMÁTICA", command=lambda: self.Chosing_Course("INTRODUÇÃO A INFORMÁTICA"))
        self.course.add_command(label="INFORMÁTICA BÁSICA (TERCEIRA IDADE)", command=lambda: self.Chosing_Course("INFORMÁTICA BÁSICA (TERCEIRA IDADE)"))
        self.course.add_command(label="PACOTE OFFICE", command=lambda: self.Chosing_Course("PACOTE OFFICE"))
        self.course.add_command(label="EXCEL(INTERMEDIÁRIO)", command=lambda: self.Chosing_Course("EXCEL(INTERMEDIÁRIO)"))
        self.course.add_command(label="EXCEL(AVANÇADO)", command=lambda: self.Chosing_Course("EXCEL(AVANÇADO)"))
        self.course.add_command(label="LÓGICA DE PROGRAMAÇÃO", command=lambda: self.Chosing_Course("INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO"))
        self.course.add_command(label="INTRODUÇÃO A ALGORÍTIMOS", command=lambda: self.Chosing_Course("INTRODUÇÃO A ALGORÍTIMOS"))
        self.course.add_command(label="INTRODUÇÃO EM ROBÓTICA", command=lambda: self.Chosing_Course("INTRODUÇÃO EM ROBÓTICA"))
        self.course.add_command(label="INTRODUÇÃO A UTILIZAÇÃO DE IAS E CHATBOTS DE FORMA PRODUTIVA", command=lambda: self.Chosing_Course("INTRODUÇÃO A UTILIZAÇÃO DE IAS E CHATBOTS DE FORMA PRODUTIVA"))
        self.course.add_command(label="INTRODUÇÃO A INOVAÇÃO E DESIGN", command=lambda: self.Chosing_Course("INTRODUÇÃO A INOVAÇÃO E DESIGN"))
        self.course.add_command(label="INSTAGRAM PARA NEGÓCIOS", command=lambda: self.Chosing_Course("INSTAGRAM PARA NEGÓCIOS"))
        self.course.add_command(label="MARKETING DIGITAL", command=lambda: self.Chosing_Course("MARKETING DIGITAL"))
        self.course.add_command(label="CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", command=lambda: self.Chosing_Course("CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS"))
        self.course.add_command(label="REDES SOCIAIS: UM GUIA PRÁTICO  PARA ALAVANCAR SUAS VENDAS", command=lambda: self.Chosing_Course("REDES SOCIAIS: UM GUIA PRÁTICO  PARA ALAVANCAR SUAS VENDAS"))
        self.course.add_command(label="SKETCHUP - SOFTWARE CRIAÇÃO DE MODELOS EM 3D", command=lambda: self.Chosing_Course("SKETCHUP - SOFTWARE CRIAÇÃO DE MODELOS EM 3D"))
        self.course.add_command(label="IMPRESSORA 3D - BÁSICO", command=lambda: self.Chosing_Course("IMPRESSORA 3D - BÁSICO"))



        self.menu_course.configure(menu=self.course)
        self.menu_course.pack(pady=5)



        # ================================================================================================================================



        # criar o menu de seleção de parceiros
        self.menu_button = ttk.Menubutton(root, text="Instituição Parceira")
        self.partner_selected = tk.StringVar()
        self.parceiro = tk.Menu(self.menu_button, tearoff=0)

        # Adiciona os parceiros aqui
        self.parceiro.add_command(label="SENAC", command=lambda: self.Chosing_Partner("SENAC"))
        self.parceiro.add_command(label="SENAI", command=lambda: self.Chosing_Partner("SENAI"))
        self.parceiro.add_command(label="SEBRAE", command=lambda: self.Chosing_Partner("SEBRAE"))

        self.menu_button.configure(menu=self.parceiro)
        self.menu_button.pack(pady=5)



        # ================================================================================================================================



        # Função para criar o menu de seleção de grupos
        self.menu_group = ttk.Menubutton(root, text="Deseja enviar mensagem por grupos?")
        self.group_selected = tk.StringVar()
        self.group = tk.Menu(self.menu_group, tearoff=0)
        self.group.add_command(label="SIM", command=lambda: self.Chosing_Group("SIM"))
        self.group.add_command(label="NÃO", command=lambda: self.Chosing_Group("NÃO"))
        self.menu_group.configure(menu=self.group)
        self.menu_group.pack(pady=5)



        # ================================================================================================================================



        self.schedule_label = ttk.Label(root, text="Horário:")
        self.schedule_label.pack()
        self.schedule_entry = ttk.Entry(self.root, width=30, bootstyle='info')
        self.schedule_entry.pack()

        self.minage_label = ttk.Label(root, text="Idade mínima:")
        self.minage_label.pack()
        self.minage_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.minage_entry.pack()

        self.duration_label = ttk.Label(root, text="Duração:")
        self.duration_label.pack()
        self.duration_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.duration_entry.pack()

        self.minrange_label = ttk.Label(root, text="De qual linha devo começar:")
        self.minrange_label.pack()
        self.minrange_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.minrange_entry.insert(0, str(self.last_line)) # Inserir o valor da última linha enviada
        self.minrange_entry.pack()

        self.maxrange_label = ttk.Label(root, text="Até qual linha devo enviar:")
        self.maxrange_label.pack()
        self.maxrange_entry = ttk.Entry(root, width=30, bootstyle='info')
        self.maxrange_entry.pack()

        # Criar um botão para enviar mensagens
        self.send_button = ttkb.Button(self.root, text="Enviar Mensagens", command=self.start_sending, bootstyle='success')
        self.send_button.pack(pady=10)

        # Criar um botão para cancelar o código
        self.cancel_button = ttkb.Button(self.root, text="Interromper código", command=self.interromper_codigo, bootstyle='danger')
        self.cancel_button.pack(pady=10)

        # CRÉDITOS
        self.credits_label = ttkb.Label(self.root, text='developed by: Lucas Ferrari', font=('Arial', 7, 'bold'))
        self.credits_label.pack(pady=5)

        self.running = False



    # ================================================================================================================================



    # Funções de suporte.
    def interromper_codigo(self):
        print("Envio encerrado")
        self.running = False

    def start_sending(self):
        if self.running:
            print("O envio de mensagens já está em andamento.")
            return
        self.running = True
        print('Iniciado o envio de mensagens!')
        t = threading.Thread(target=self.send_messages)
        t.start()

    def Chosing_Course(self, curso: str):
        self.menu_course.config(text=curso)
        self.course_selected.set(curso)

    def Chosing_Partner(self, parceiro: str):
        self.menu_button.config(text=parceiro)
        self.partner_selected.set(parceiro)

    def Chosing_Group(self, grupo: str):
        self.menu_group.config(text=grupo)
        self.group_selected.set(grupo)

    def change_theme(self, theme):
        # Atualiza o tema na interface
        self.style.theme_use(theme)
        # Salva o tema escolhido nas configurações
        self.settings["theme"] = theme
        save_settings(self.settings)



    def encontra_categoria(self, curso_de_envio: str) -> list:
        # DEFINE AS CATEGORIAS
        Dev = ["INTRODUÇÃO A LÓGICA DE PROGRAMAÇÃO", "INTRODUÇÃO A ALGORÍTIMOS", "INTRODUÇÃO EM ROBÓTICA", "INTRODUÇÃO A UTILIZAÇÃO DE IAS E CHATBOTS DE FORMA PRODUTIVA"]
        Marketing = ["INSTAGRAM PARA NEGÓCIOS", "MARKETING DIGITAL", "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", "REDES SOCIAIS: UM GUIA PRÁTICO  PARA ALAVANCAR SUAS VENDAS"]
        Design = ["INTRODUÇÃO A INOVAÇÃO E DESIGN", "CRIAÇÃO DE MÍDIAS PARA REDES SOCIAIS", "SKETCHUP - SOFTWARE CRIAÇÃO DE MODELOS EM 3D", "IMPRESSORA 3D - BÁSICO"]
        Pacote_Office = ["PACOTE OFFICE", "EXCEL(INTERMEDIÁRIO)", "EXCEL(AVANÇADO)"]
        Basicos = ["INTRODUÇÃO A INFORMÁTICA", "INFORMÁTICA BÁSICA (TERCEIRA IDADE)"]

        lista_categoria = [Dev, Marketing, Design, Pacote_Office, Basicos]

        for Categorias in lista_categoria:
            if curso_de_envio in Categorias:
                return Categorias



    # ================================================================================================================================



    def send_messages(self):
        # Salva as entradas obtidas na interface.
        curso_de_envio = self.course_selected.get()
        parceiro = self.partner_selected.get()
        horario_do_curso = self.schedule_entry.get()
        data_de_duracao = self.duration_entry.get()
        linhamin = int(self.minrange_entry.get())
        linhamax = int(self.maxrange_entry.get())
        idademin = int(self.minage_entry.get())
        por_grupo = self.group_selected.get()

        # Ler o arquivo Excel
        alunos = pd.read_excel('alunos.xlsx')

        # Conjunto para armazenar números de telefone já processados
        numeros_enviados = set()

        # Definir running como True para iniciar o envio de mensagens
        self.running = True

        total_alunos = linhamax - linhamin
        ultima_linha_enviada = None

        for x in range(linhamin, linhamax):
            if not self.running:
                print('Código interrompido na linha: {0}'.format(x))
                break

            cursos = alunos.loc[x, "Dentre as opções qual curso gostaria de fazer?"]
            lista_cursos = cursos.split(sep=', ')

            if por_grupo == "SIM":
                categoria = self.encontra_categoria(curso_de_envio)
                for curso in lista_cursos:
                    if curso in categoria:
                        nome = alunos.loc[x, 'Nome Completo']
                        telefone = int(alunos.loc[x, "Whatsapp com DDD (somente números - sem espaço)"])

                        if telefone in numeros_enviados:
                            continue

                        mensagem = f"Olá *{nome}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\n Nós iremos iniciar em parceria com o *{parceiro}*, o curso:\n\n 🌟*{curso_de_envio}*.🌟 \n\n Todos podem participar desde que sejam maior de *{idademin}* anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.🎓\n\n🎯 Duração do curso: *{data_de_duracao}*\n\n 🕒 Horário: *{horario_do_curso}* \n\n⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá🗺️ -*\n\n*🏫 MODALIDADE: curso é PRESENCIAL E 100% GRATUITO! 🎉* \n\n Qualquer dúvida, estamos à disposição! Esperamos você! 😉"

                        link_mensagem_whatsapp = f'https://web.whatsapp.com/send/?phone={telefone}&text={quote(mensagem)}'
                        webbrowser.open(link_mensagem_whatsapp)
                        sleep(6)
                        sleep(4)
                        pyautogui.press('enter')
                        sleep(4)
                        pyautogui.hotkey('ctrl', 'w')
                        numeros_enviados.add(telefone)
                        ultima_linha_enviada = x

            else:
                for curso in lista_cursos:
                    if curso.upper() == curso_de_envio.upper():
                        nome = alunos.loc[x, 'Nome Completo']
                        telefone = int(alunos.loc[x, 'Whatsapp com DDD (somente números - sem espaço)'])

                        if telefone in numeros_enviados:
                            continue

                        mensagem = f"Olá *{nome}.* Nós somos da AMTECH - Agência Maringá de Tecnologia e Inovação. entramos em contato porque você demonstrou interesse em cursos de tecnologia preenchendo um formulário.📋\n\n Nós iremos iniciar em parceria com o *{parceiro}*, o curso:\n\n 🌟*{curso_de_envio}*.🌟 \n\n Todos podem participar desde que sejam maior de *{idademin}* anos e tenham a escolaridade mínima 5º ano do Ensino Fundamental.🎓\n\n🎯 Duração do curso: *{data_de_duracao}*\n\n 🕒 Horário: *{horario_do_curso}* \n\n⚠️ Atenção: As vagas são limitadas! Responda o mais rápido possível! 🏃‍♂️💨 📢*\n\n*📍Local: Acesso 1 | Piso Superior Terminal Urbano - Av. Tamandaré, 600 - Zona 01, Maringá🗺️  -*\n\n*🏫 MODALIDADE: curso é PRESENCIAL E 100% GRATUITO! 🎉* \n\nQualquer dúvida, estamos à disposição! Esperamos você! 😉"

                        link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                        webbrowser.open(link_mensagem_whatsapp)
                        sleep(6)
                        sleep(4)
                        pyautogui.press('enter')
                        sleep(4)
                        pyautogui.hotkey('ctrl', 'w')
                        numeros_enviados.add(telefone)
                        ultima_linha_enviada = x

                        # Aqui é onde o log é adicionado, após o envio da mensagem
                        logging.info(f'Messagem enviada para: {nome}, Telefone: {telefone}, Curso: {curso_de_envio}')
                        logging.info(f'Última linha enviada: {ultima_linha_enviada}')  # Adiciona ao log

                    if ultima_linha_enviada is not None:
                        print(f'Ultima linha enviada: {ultima_linha_enviada}')
                        logging.info(f'Ultima linha enviada: {ultima_linha_enviada}') # Adiciona ao log
                        save_last_line(ultima_linha_enviada)  # Salvar a última linha enviada
                        


        print("Todas as linhas foram lidas!")
        self.running = False

        ttk.Label(self.root, text="Envio de mensagens concluído!", foreground="green").pack(pady=10)



# Criar a GUI
root = tk.Tk()
gui = CourseOfferGUI(root)
root.mainloop()


