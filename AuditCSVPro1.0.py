import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import os
import shutil
import re

class AuditCSVPro(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações de Interface
        self.title("Audit CSV Pro 1.0")
        self.width = 450
        self.height = 640
        self.centralizar_janela()
        self.resizable(False, False)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Cabeçalho
        ctk.CTkLabel(self, text="Audit CSV Pro 1.0", font=("Roboto", 22, "bold")).pack(pady=15)

        # Container Auditor
        self.frame_auditor = ctk.CTkFrame(self)
        self.frame_auditor.pack(pady=5, padx=25, fill="x")
        ctk.CTkLabel(self.frame_auditor, text="DADOS DO AUDITOR", font=("Roboto", 11, "bold")).pack(pady=5)

        self.ent_nome = ctk.CTkEntry(self.frame_auditor, placeholder_text="Nome Completo", height=30)
        self.ent_nome.pack(pady=5, padx=15, fill="x")

        self.ent_telefone = ctk.CTkEntry(self.frame_auditor, placeholder_text="Telefone: (99) 99999-9999", height=30)
        self.ent_telefone.pack(pady=5, padx=15, fill="x")
        self.ent_telefone.bind("<KeyRelease>", self.formatar_telefone_manual)

        self.ent_email = ctk.CTkEntry(self.frame_auditor, placeholder_text="E-mail Corporativo", height=30)
        self.ent_email.pack(pady=5, padx=15, fill="x")

        self.ent_empresa = ctk.CTkEntry(self.frame_auditor, placeholder_text="Empresa (Opcional)", height=30)
        self.ent_empresa.pack(pady=5, padx=15, fill="x")

        # Container Arquivo
        self.frame_file = ctk.CTkFrame(self)
        self.frame_file.pack(pady=10, padx=25, fill="x")
        self.btn_select = ctk.CTkButton(self.frame_file, text="Selecionar CSV", command=self.select_file, height=32)
        self.btn_select.pack(pady=10)
        self.label_path = ctk.CTkLabel(self.frame_file, text="Nenhum arquivo selecionado", font=("Roboto", 10), text_color="gray")
        self.label_path.pack(pady=2)

        # Botão Ação
        self.btn_run = ctk.CTkButton(self, text="GERAR RELATÓRIO COMPLETO", 
                                     command=self.run_audit, fg_color="#27ae60", hover_color="#219150", 
                                     height=40, font=("Roboto", 13, "bold"))
        self.btn_run.pack(pady=20)

        ctk.CTkLabel(self, text="Desenvolvido por: Moises Costa - 31 99954-4458", 
                     font=("Roboto", 12, "bold"), text_color="#ABB2B9").pack(side="bottom", pady=15)

        self.file_path = ""

    def centralizar_janela(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (self.width // 2)
        y = (screen_height // 2) - (self.height // 2)
        self.geometry(f"{self.width}x{self.height}+{x}+{y}")

    def formatar_telefone_manual(self, event):
        if event.keysym in ("BackSpace", "Delete", "Left", "Right", "Tab"): return
        v = self.ent_telefone.get()
        nums = "".join(filter(str.isdigit, v))[:11]
        formato = ""
        if len(nums) > 0:
            formato = "(" + nums[:2]
            if len(nums) > 2:
                formato += ") " + nums[2:7]
                if len(nums) > 7:
                    formato += "-" + nums[7:]
        if v != formato:
            self.ent_telefone.delete(0, "end")
            self.ent_telefone.insert(0, formato)

    def select_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if self.file_path:
            nome = os.path.basename(self.file_path)
            self.label_path.configure(text=f"Arquivo: {nome[:25]}...", text_color="#3498db")

    def run_audit(self):
        nome, tel, email = self.ent_nome.get().strip(), self.ent_telefone.get().strip(), self.ent_email.get().strip()
        empresa = self.ent_empresa.get().strip() or "N/A"

        if not nome or len(tel) < 14 or not email or not self.file_path:
            messagebox.showwarning("Atenção", "Preencha todos os campos e selecione o arquivo.")
            return

        try:
            # 1. Leitura do Arquivo
            df = pd.read_csv(self.file_path, sep=None, engine='python', on_bad_lines='skip', encoding='utf-8-sig')
            
            # 2. Pasta de Saída
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            folder_name = datetime.now().strftime("Auditoria_%Y%m%d_%H%M%S")
            out_dir = os.path.join(desktop, "Relatorios_Auditoria_Pro", folder_name)
            os.makedirs(out_dir, exist_ok=True)
            
            # 3. Auditoria Detalhada de Erros
            error_data = [["Linha", "Coluna / Campo", "Tipo de Erro Identificado"]]
            
            for r_idx, row in df.iterrows():
                for c_idx, value in enumerate(df.columns):
                    celula = row[value]
                    tipo_erro = None
                    
                    # Identificação de Tipos de Erro
                    if pd.isnull(celula):
                        tipo_erro = "Campo Vazio (Falta de Informação)"
                    elif str(celula).strip() == "":
                        tipo_erro = "Preenchido apenas com Espaços"
                    elif str(celula).lower() in ['null', 'none', 'nan', 'undefined']:
                        tipo_erro = "Texto reservado inválido (Null/None)"
                    
                    if tipo_erro:
                        error_data.append([str(r_idx + 2), str(value), tipo_erro])

            # 4. Geração do PDF Relatório
            pdf_path = os.path.join(out_dir, "Relatorio_Auditoria_Completo.pdf")
            doc = SimpleDocTemplate(pdf_path, pagesize=letter)
            styles = getSampleStyleSheet()
            elementos = []

            elementos.append(Paragraph("RELATÓRIO DE AUDITORIA E DIAGNÓSTICO", styles['Title']))
            elementos.append(Spacer(1, 12))
            
            resumo = [
                f"<b>Auditor Responsável:</b> {nome}",
                f"<b>Contato:</b> {tel} | {email}",
                f"<b>Data da Execução:</b> {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                f"<b>Base Analisada:</b> {os.path.basename(self.file_path)}",
                f"<b>Total de Linhas:</b> {len(df)} | <b>Total de Inconformidades:</b> {len(error_data)-1}"
            ]
            for r in resumo: elementos.append(Paragraph(r, styles['Normal']))
            
            elementos.append(Spacer(1, 20))
            elementos.append(Paragraph("Detalhamento Técnico dos Erros:", styles['Heading3']))
            elementos.append(Spacer(1, 10))

            if len(error_data) > 1:
                t = Table(error_data, colWidths=[50, 150, 250])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.whitesmoke, colors.lightgrey])
                ]))
                elementos.append(t)
            else:
                elementos.append(Paragraph("✅ Nenhuma inconformidade encontrada nos critérios de campos nulos.", styles['Normal']))

            elementos.append(Spacer(1, 30))
            elementos.append(Paragraph("<font size='8'>Gerado por Audit CSV Pro 1.0 - Desenvolvido por Moises Costa - 31 99954-4458 - moisescarvalho33@gmail.com</font>", styles['Normal']))
            doc.build(elementos)

            # 5. Exportação Excel e Cópia Original
            pd.DataFrame(error_data[1:], columns=error_data[0]).to_excel(os.path.join(out_dir, "mapeamento_erros.xlsx"), index=False)
            shutil.copy(self.file_path, os.path.join(out_dir, "BASE_ORIGINAL.csv"))

            messagebox.showinfo("Concluído", f"Auditoria finalizada com sucesso!\nPasta: {folder_name}")

        except Exception as e:
            messagebox.showerror("Erro", f"Falha técnica no processamento: {e}")

if __name__ == "__main__":
    app = AuditCSVPro()
    app.mainloop()