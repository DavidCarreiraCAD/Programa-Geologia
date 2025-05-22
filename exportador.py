# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import shutil
import os
import re
import subprocess
import sys

# Função para aceder a ficheiros no modo .exe ou .py
def get_resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# Caminho base
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(__file__)

ARQUIVOS = {
    "ZZZ_ProspBlocks.dwg": get_resource_path(os.path.join("resources", "ZZZ_ProspBlocks.dwg")),
    "run_sondagens.scr": get_resource_path(os.path.join("resources", "run_sondagens.scr")),
    "ImportLogsPlanta.lsp": get_resource_path(os.path.join("resources", "ImportLogsPlanta.lsp")),
    "run_logsplanta.scr": get_resource_path(os.path.join("resources", "run_logsplanta.scr")),
    "LogTransformMultiple.lsp": get_resource_path(os.path.join("resources", "LogTransformMultiple.lsp"))

}

def encontrar_acad():
    versoes = ["2026", "2025", "2024", "2023"]
    base_path = r"C:\Program Files\Autodesk\AutoCAD {}"

    for versao in versoes:
        acad_path = os.path.join(base_path.format(versao), "acad.exe")
        if os.path.isfile(acad_path):
            return acad_path

    return None

ACAD_PATH = encontrar_acad()

if not ACAD_PATH:
    messagebox.showerror("Erro", "Não foi encontrada nenhuma instalação do AutoCAD nas versões 2023 a 2026.")
    sys.exit(1)

class MainMenu:
    def __init__(self, master):
        self.master = master
        master.title("Seleciona o Comando")
        master.geometry("350x300")
        master.configure(bg="#f2f2f2")
        master.resizable(False, False)

        frame = ttk.Frame(master, padding=20)
        frame.pack(expand=True)

        # Título estilizado
        titulo = ttk.Label(frame, text="Seleciona o Comando", font=("Segoe UI", 14, "bold"))
        titulo.pack(pady=(0, 20))

        # Estilo de botões uniforme
        estilo = ttk.Style()
        estilo.configure("TButton", font=("Segoe UI", 11), padding=10)

        # Botões
        btn_logtransform = ttk.Button(frame, text="LogTransform", command=self.abrir_logtransform)
        btn_logtransform.pack(fill="x", pady=5)

        btn_importlogs = ttk.Button(frame, text="ImportLogsPlanta", command=self.import_logs_planta)
        btn_importlogs.pack(fill="x", pady=5)

        btn_cancelar = ttk.Button(frame, text="Cancelar", command=master.quit)
        btn_cancelar.pack(fill="x", pady=5)

    def abrir_logtransform(self):
        self.master.destroy()
        root_log = tk.Tk()
        app_log = LogTransform(root_log)
        root_log.mainloop()

    def import_logs_planta(self):
        self.master.destroy()
        root_import = tk.Tk()
        import_app = ImportLogsPlanta(root_import)
        root_import.mainloop()

class ImportLogsPlanta:
    def __init__(self, master):
        self.master = master
        master.title("ImportLogsPlanta")
        master.geometry("500x140")
        master.resizable(False, False)

        self.ficheiro_dwg = ""
        self.pasta = ""

        titulo = ttk.Label(master, text="ImportLogsPlanta", font=("Segoe UI", 16, "bold"))
        titulo.pack(pady=(10, 5))

        frame = ttk.Frame(master, padding=20)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Ficheiro DWG:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.dwg_entry = ttk.Entry(frame, width=45, state="readonly")
        self.dwg_entry.grid(row=0, column=1, padx=5, pady=5)
        self.btn_browser = ttk.Button(frame, text="Procurar", command=self.selecionar_ficheiro_dwg)
        self.btn_browser.grid(row=0, column=2, padx=5, pady=5)

        self.btn_gerar = ttk.Button(frame, text="Gerar", command=self.gerar)
        self.btn_gerar.grid(row=1, column=0, padx=5, pady=10)

        self.btn_cancelar = ttk.Button(frame, text="Cancelar", command=self.master.quit)
        self.btn_cancelar.grid(row=1, column=1, padx=5, pady=10)

        self.btn_concluir = ttk.Button(frame, text="Concluir", command=self.limpar_arquivos, state="disabled")
        self.btn_concluir.grid(row=1, column=2, padx=5, pady=10)

    def selecionar_ficheiro_dwg(self):
        ficheiro = filedialog.askopenfilename(filetypes=[("Ficheiros DWG", "*.dwg")])
        if ficheiro:
            self.ficheiro_dwg = ficheiro
            self.pasta = os.path.dirname(ficheiro)
            self.dwg_entry.config(state="normal")
            self.dwg_entry.delete(0, tk.END)
            self.dwg_entry.insert(0, self.ficheiro_dwg)
            self.dwg_entry.config(state="readonly")

    def gerar(self):
        if not self.ficheiro_dwg or not os.path.isfile(self.ficheiro_dwg):
            messagebox.showerror("Erro", "Seleciona um ficheiro DWG válido primeiro!")
            return

        try:
            for nome in ["ZZZ_ProspBlocks.dwg", "run_logsplanta.scr", "ImportLogsPlanta.lsp"]:
                shutil.copyfile(ARQUIVOS[nome], os.path.join(self.pasta, nome))

            ficheiros_csv = [f for f in os.listdir(self.pasta) if f.lower().endswith(".csv")]
            if not ficheiros_csv:
                messagebox.showerror("Erro", "Nenhum ficheiro .csv encontrado na pasta do DWG.")
                return

            caminho_lsp = os.path.join(self.pasta, "ImportLogsPlanta.lsp")
            caminho_dwg_bloques = os.path.join(self.pasta, "ZZZ_ProspBlocks.dwg").replace("\\", "/")
            caminho_csv = os.path.join(self.pasta, ficheiros_csv[0]).replace("\\", "/")

            with open(caminho_lsp, "r", encoding="utf-8") as f:
                conteudo_lsp = f.read()

            conteudo_lsp = re.sub(r'\(setq ficheiroBlocos\s+"[^"]+"\)', f'(setq ficheiroBlocos "{caminho_dwg_bloques}")', conteudo_lsp)
            conteudo_lsp = re.sub(r'\(setq ficheiroCsv\s+"[^"]+"\)', f'(setq ficheiroCsv "{caminho_csv}")', conteudo_lsp)

            with open(caminho_lsp, "w", encoding="utf-8") as f:
                f.write(conteudo_lsp)

            caminho_scr = os.path.join(self.pasta, "run_logsplanta.scr")
            with open(caminho_scr, "r", encoding="utf-8") as f:
                conteudo_scr = f.read()

            caminho_lsp_escaped = caminho_lsp.replace("\\", "/")
            conteudo_scr = re.sub(r'\(load\s+"[^"]*ImportLogsPlanta\.lsp"\)', f'(load "{caminho_lsp_escaped}")', conteudo_scr)

            with open(caminho_scr, "w", encoding="utf-8") as f:
                f.write(conteudo_scr)

            # Executa o AutoCAD com o DWG selecionado e o script
            subprocess.Popen([ACAD_PATH, self.ficheiro_dwg, "/b", caminho_scr])

            self.btn_concluir.config(state="normal")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

    def limpar_arquivos(self):
        try:
            for nome in ["ImportLogsPlanta.lsp", "run_logsplanta.scr", "ZZZ_ProspBlocks.dwg"]:
                caminho = os.path.join(self.pasta, nome)
                if os.path.exists(caminho):
                    os.remove(caminho)
            messagebox.showinfo("Limpeza", "Processo concluído.")
            self.master.quit()
            self.master.destroy()
        except Exception as e:
            messagebox.showwarning("Aviso", f"Erro ao eliminar ficheiros: {e}")
            self.btn_concluir.config(state="disabled")



class LogTransform:
    def __init__(self, master):
        self.master = master
        master.title("LogTransform")
        master.geometry("500x160")
        master.resizable(False, False)

        self.pasta = ""

        titulo = ttk.Label(master, text="LogTransform", font=("Segoe UI", 16, "bold"))
        titulo.pack(pady=(10, 5))

        frame = ttk.Frame(master, padding=20)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Pasta:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.pasta_entry = ttk.Entry(frame, width=45, state="readonly")
        self.pasta_entry.grid(row=0, column=1, padx=5, pady=5)
        self.btn_browser = ttk.Button(frame, text="Procurar", command=self.selecionar_pasta)
        self.btn_browser.grid(row=0, column=2, padx=5, pady=5)

        frame_escala = ttk.Frame(frame)
        frame_escala.grid(row=1, column=0, columnspan=3, sticky="ew", pady=10)

        for i in range(5):
            frame_escala.columnconfigure(i, weight=1)

        ttk.Label(frame_escala, text="Escala:").grid(row=0, column=0, sticky="e", padx=5)
        self.escala_entry = ttk.Entry(frame_escala, width=10)
        self.escala_entry.grid(row=0, column=1, sticky="ew", padx=5)

        self.btn_gerar = ttk.Button(frame_escala, text="Gerar", command=self.gerar)
        self.btn_gerar.grid(row=0, column=2, sticky="ew", padx=5)

        self.btn_cancelar = ttk.Button(frame_escala, text="Cancelar", command=self.master.quit)
        self.btn_cancelar.grid(row=0, column=3, sticky="ew", padx=5)

        self.btn_concluir = ttk.Button(frame_escala, text="Concluir", command=self.limpar_arquivos, state="disabled")
        self.btn_concluir.grid(row=0, column=4, sticky="ew", padx=5)

    def selecionar_pasta(self):
        pasta_selecionada = filedialog.askdirectory()
        if pasta_selecionada:
            self.pasta = pasta_selecionada
            self.pasta_entry.config(state="normal")
            self.pasta_entry.delete(0, tk.END)
            self.pasta_entry.insert(0, self.pasta)
            self.pasta_entry.config(state="readonly")

    def gerar(self):
        if not self.pasta:
            messagebox.showerror("Erro", "Seleciona uma pasta primeiro!")
            return

        escala = self.escala_entry.get()
        if not escala:
            messagebox.showerror("Erro", "Preenche o campo da escala!")
            return

        try:
            caminho_destino = self.pasta
            for nome, origem in ARQUIVOS.items():
                shutil.copyfile(origem, os.path.join(caminho_destino, nome))

            caminho_lsp_novo = os.path.join(caminho_destino, "LogTransformMultiple.lsp")
            with open(ARQUIVOS["LogTransformMultiple.lsp"], "r", encoding="utf-8") as f:
                conteudo_lsp = f.read()

            novo_dwg = os.path.join(caminho_destino, "ZZZ_ProspBlocks.dwg").replace("\\", "/")
            novo_csv = os.path.join(caminho_destino, "FormatoLogs.csv").replace("\\", "/")
            novo_dwg_externo = r"C:\Users\DPC\OneDrive - COBAGroup\Desktop\REVIT_David\SONDAGENS_DWG_TESTES\ProspBlocksTeste.dwg".replace("\\", "/")

            conteudo_lsp = re.sub(r'\(setq ficheiro\s+"[^"]+"\)', f'(setq ficheiro "{novo_dwg}")', conteudo_lsp)
            conteudo_lsp = re.sub(r'\(setq file\s+\(open\s+"[^"]+"\s+"r"\)\)', f'(setq file (open "{novo_csv}" "r"))', conteudo_lsp)
            conteudo_lsp = re.sub(r'\(setq escala\s+[^\)]+\)', f'(setq escala {escala})', conteudo_lsp)
            conteudo_lsp = re.sub(r'\(setq prospBlocksTeste\s+"[^"]+"\)', f'(setq prospBlocksTeste "{novo_dwg_externo}")', conteudo_lsp)

            with open(caminho_lsp_novo, "w", encoding="utf-8") as f:
                f.write(conteudo_lsp)

            caminho_scr_novo = os.path.join(caminho_destino, "run_sondagens.scr")
            with open(ARQUIVOS["run_sondagens.scr"], "r", encoding="utf-8") as f:
                conteudo_scr = f.read()

            novo_lsp_path = caminho_lsp_novo.replace("\\", "/")
            conteudo_scr = re.sub(r'\(load\s+"[^"]*LogTransformMultiple\.lsp"\)', f'(load "{novo_lsp_path}")', conteudo_scr)

            with open(caminho_scr_novo, "w", encoding="utf-8") as f:
                f.write(conteudo_scr)

            for file in os.listdir(caminho_destino):
                if file.lower().endswith(".dxf"):
                    caminho_dxf = os.path.join(caminho_destino, file)
                    caminho_scr = os.path.join(caminho_destino, "run_sondagens.scr")
                    subprocess.Popen([
                        ACAD_PATH,
                        caminho_dxf,
                        "/b",
                        caminho_scr
                    ])

            self.btn_concluir.grid()

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

        self.btn_concluir.config(state="normal")

    def limpar_arquivos(self):
        try:
            for nome in ["LogTransformMultiple.lsp", "run_sondagens.scr", "ZZZ_ProspBlocks.dwg"]:
                caminho = os.path.join(self.pasta, nome)
                if os.path.exists(caminho):
                    os.remove(caminho)
            messagebox.showinfo("Limpeza", "Processo concluído.")
            self.master.quit()
            self.master.destroy()
        except Exception as e:
            messagebox.showwarning("Aviso", f"Erro ao eliminar ficheiros: {e}")

            self.btn_concluir.config(state="disabled")



if __name__ == "__main__":
    root = tk.Tk()
    main_menu = MainMenu(root)
    root.mainloop()
