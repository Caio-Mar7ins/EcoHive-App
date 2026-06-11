# EcoHive-App
#Programa realizado para fins acadêmicos. Utilizei somente Python para o desenvolvimento deste sistema.
import customtkinter as ctk
from PIL import Image
from tkinter import END
import pandas as pd
import os
from datetime import datetime
from zoneinfo import ZoneInfo
import textwrap


class BackEnd():

    EXCEL_PATH = 'ecohive_.xlsx'
    EXCEL_USERS_SHEET = 'usuarios'
    EXCEL_DESCARTE_SHEET = 'descarte'
    EXCEL_COLUMNS = ['id', 'usuario', 'bloco',
                     'apartamento', 'email', 'senha', 'pontuacao']

    def criar_excel(self):

        if not os.path.exists(self.EXCEL_PATH):
            df_users = pd.DataFrame(columns=self.EXCEL_COLUMNS)

            if 'pontuacao' in df_users.columns:
                df_users['pontuacao'] = 0
            df_descarte = pd.DataFrame(columns=[
                'data', 'horario', 'usuario', 'bloco', 'apartamento', 'tipo_lixo', 'peso_kg'
            ])
            with pd.ExcelWriter(self.EXCEL_PATH, engine='openpyxl') as writer:
                df_users.to_excel(
                    writer, sheet_name=self.EXCEL_USERS_SHEET, index=False)
                df_descarte.to_excel(
                    writer, sheet_name=self.EXCEL_DESCARTE_SHEET, index=False)

    def ler_usuarios_df(self):

        self.criar_excel()
        try:
            df = pd.read_excel(
                self.EXCEL_PATH, sheet_name=self.EXCEL_USERS_SHEET)
            for c in self.EXCEL_COLUMNS:
                if c not in df.columns:
                    df[c] = pd.NA
            df = df[self.EXCEL_COLUMNS]
            return df
        except Exception as e:
            print("Erro ao ler sheet usuarios:", e)
            return pd.DataFrame(columns=self.EXCEL_COLUMNS)

    def gravar_usuarios_df(self, df: pd.DataFrame):

        for c in self.EXCEL_COLUMNS:
            if c not in df.columns:
                df[c] = pd.NA
        df = df[self.EXCEL_COLUMNS]

        try:
            df_descarte = pd.read_excel(
                self.EXCEL_PATH, sheet_name=self.EXCEL_DESCARTE_SHEET)
        except Exception:
            df_descarte = pd.DataFrame(columns=[
                                       'data', 'horario', 'usuario', 'bloco', 'apartamento', 'tipo_lixo', 'peso_kg'])

        with pd.ExcelWriter(self.EXCEL_PATH, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=self.EXCEL_USERS_SHEET, index=False)
            df_descarte.to_excel(
                writer, sheet_name=self.EXCEL_DESCARTE_SHEET, index=False)

    def inserir_usuario_excel(self, usuario, bloco, apartamento, email, senha):

        df = self.ler_usuarios_df()

        if 'id' in df.columns and not df.empty:
            try:
                max_id = pd.to_numeric(df['id'], errors='coerce').dropna()
                novo_id = int(max_id.max()) + 1 if not max_id.empty else 1
            except Exception:
                novo_id = 1
        else:
            novo_id = 1

        nova_linha = {
            'id': novo_id,
            'usuario': usuario,
            'bloco': bloco if bloco != '' else None,
            'apartamento': int(apartamento) if (apartamento is not None and apartamento != '') else None,
            'email': email,
            'senha': senha,
            'pontuacao': 0
        }

        novo_df = pd.DataFrame([nova_linha])
        df = pd.concat([df, novo_df], ignore_index=True)

        self.gravar_usuarios_df(df)

    def ler_descarte_df(self):

        self.criar_excel()
        try:
            df = pd.read_excel(
                self.EXCEL_PATH, sheet_name=self.EXCEL_DESCARTE_SHEET)

            expected = ['data', 'horario', 'usuario', 'bloco',
                        'apartamento', 'tipo_lixo', 'peso_kg']
            for c in expected:
                if c not in df.columns:
                    df[c] = pd.NA
            return df[expected]
        except Exception as e:
            print("Erro ao ler sheet descarte:", e)
            return pd.DataFrame(columns=['data', 'horario', 'usuario', 'bloco', 'apartamento', 'tipo_lixo', 'peso_kg'])

    def gravar_descarte(self, tipo_lixo: str, peso_kg: float, usuario: str = None, bloco: str = None, apartamento=None):

        try:
            df_desc = self.ler_descarte_df()
            nova_linha = {
                'data': datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%Y-%m-%d'),
                'horario': datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%H:%M:%S'),
                'usuario': usuario or '',
                'bloco': bloco or '',
                'apartamento': apartamento if apartamento is not None else '',
                'tipo_lixo': tipo_lixo,
                'peso_kg': float(peso_kg)
            }
            novo_df = pd.DataFrame([nova_linha])
            df_desc = pd.concat([df_desc, novo_df], ignore_index=True)

            df_users = self.ler_usuarios_df()

            with pd.ExcelWriter(self.EXCEL_PATH, engine='openpyxl') as writer:
                df_users.to_excel(
                    writer, sheet_name=self.EXCEL_USERS_SHEET, index=False)
                df_desc.to_excel(
                    writer, sheet_name=self.EXCEL_DESCARTE_SHEET, index=False)

            return True
        except Exception as e:
            print("Erro ao gravar descarte:", e)
            return False

    def somatorio_por_tipo_por_bloco(self, bloco: str):

        try:
            df_desc = self.ler_descarte_df()
            if df_desc.empty:
                return {}

            df_filtrado = df_desc[df_desc['bloco'].astype(
                str).str.strip().str.lower() == str(bloco).strip().lower()]
            if df_filtrado.empty:
                return {}

            df_filtrado = df_filtrado.copy()
            df_filtrado['peso_kg'] = pd.to_numeric(
                df_filtrado['peso_kg'], errors='coerce').fillna(0.0)
            soma = df_filtrado.groupby('tipo_lixo', as_index=False)[
                'peso_kg'].sum()
            result = {row['tipo_lixo']: float(
                row['peso_kg']) for _, row in soma.iterrows()}
            return result
        except Exception as e:
            print("Erro ao calcular somatório:", e)
            return {}

    def salvar_pontuacao_usuario(self, user_identifier=None):

        try:
            df = self.ler_usuarios_df()
            if df.empty:
                return False

            uid = None
            uname = None
            if user_identifier is not None:
                try:
                    uid = int(user_identifier)
                except Exception:
                    uname = str(user_identifier)
            else:
                if getattr(self, 'current_user', None):
                    uid = self.current_user.get('id')
                    uname = self.current_user.get('usuario')

            if uid is not None and uid != '':
                mask = pd.to_numeric(df['id'], errors='coerce') == int(uid)
            elif uname:
                mask = df['usuario'].astype(str).str.strip(
                ).str.lower() == str(uname).strip().lower()
            else:
                return False

            if not mask.any():
                return False

            pontos = float(getattr(self, 'pontuacao_atual', 0) or 0)

            df.loc[mask, 'pontuacao'] = pontos

            self.gravar_usuarios_df(df)
            return True
        except Exception as e:
            print("Erro ao salvar pontuação no Excel:", e)
            return False


class EcoHiveApp(ctk.CTk, BackEnd):
    _PLACEHOLDER_COLOR = "#9E9E9E"
    _TEXT_COLOR = "#FFFFFF"

    def __init__(self):
        super().__init__()

        self.criar_excel()

        self._bg_scale_x = 1.5
        self._bg_scale_y = 1.3

        self._default_bg = "#FFFFFF"

        self.frame_login = None
        self.frame_cadastro = None
        self.frame_home = None
        self.frame_resgate = None

        self.current_user = None

        self.fundo_img_original = None
        self.nome_app_img_original = None
        self.fundo_ctkimage = None
        self.nome_app_ctkimage = None
        self.fundo_img = None
        self.nome_app_img = None
        self.label_fundo = None
        self.label_nome_app = None

        self.var_ver_senha_login = ctk.BooleanVar(value=False)
        self.var_ver_senha_cadastro = ctk.BooleanVar(value=False)

        self.tipos_lixo_list = ['Plástico', 'Papel',
                                'Vidro', 'Lata', 'Metal', 'Eletronico']

        self.pontos_por_tipo = {'Papel': 5, 'Plástico': 10,
                                'Vidro': 3, 'Lata': 50, 'Metal': 15, 'Eletronico': 30}

        self.pontuacao_atual = 0

        self.somatorio_labels = {}

        self.recompensas = [
            {"nome": "5% off na taxa condominial", "custo": 1000},
            {"nome": "Vaga prioritaria", "custo": 2000},
            {"nome": " 10% off no mercadinho e taxa condominial", "custo": 4750},
            {"nome": 'Premio Morador Verde + R$300,00 em voucher', "custo": 7000},
        ]

        self.configuracoes_janela_principal()
        self.elementos_fixos()
        self.criar_frames_views()
        self.mostrar_tela_login()

        self.bind('<Configure>', self._on_relize)

    def configuracoes_janela_principal(self):
        self.geometry('350x600')
        self.title('EcoHive')
        self.resizable(False, False)
        self.configure(fg_color=self._default_bg)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

    def elementos_fixos(self):
        try:
            self.fundo_img_original = Image.open(
                'EcoHive_simbolo.png').convert('RGBA')
            cur_w = self.winfo_width() if self.winfo_width() > 1 else 350
            cur_h = self.winfo_height() if self.winfo_height() > 1 else 600
            target_w = max(1, int(cur_w * self._bg_scale_x))
            target_h = max(1, int(cur_h * self._bg_scale_y))
            bg_red = self.fundo_img_original.resize(
                (target_w, target_h), Image.LANCZOS)
            self.fundo_ctkimage = ctk.CTkImage(
                bg_red, size=(target_w, target_h))
            self.label_fundo = ctk.CTkLabel(
                self, text='', image=self.fundo_ctkimage, fg_color='transparent')
            # usa place (mantido)
            self.label_fundo.place(relx=0.5, rely=0.5, anchor='center')

            self.nome_app_img_original = Image.open(
                'EcoHive_1.png').convert('RGBA')
            logo_resized = self.nome_app_img_original.resize(
                (250, 250), Image.LANCZOS)
            self.nome_app_ctkimage = ctk.CTkImage(
                logo_resized, size=(250, 250))
            self.label_nome_app = ctk.CTkLabel(
                self, text='', image=self.nome_app_ctkimage, fg_color='#fbfaf6', bg_color='#fbfaf6')
            self.label_nome_app.place(relx=0.54, rely=0.185, anchor='center')

        except FileNotFoundError as e:
            print("Imagens não encontradas, usando cor de fundo. Erro:", e)
            self.configure(fg_color=self._default_bg)

    def _on_relize(self, event):
        if event.widget != self or event.width <= 1 or event.height <= 1:
            return
        if hasattr(self, 'fundo_img_original') and self.fundo_img_original and getattr(self, 'label_fundo', None):
            try:
                bg_w = max(1, int(event.width * self._bg_scale_x))
                bg_h = max(1, int(event.height * self._bg_scale_y))
                bg_red = self.fundo_img_original.resize(
                    (bg_w, bg_h), Image.LANCZOS)
                self.fundo_ctkimage = ctk.CTkImage(bg_red, size=(bg_w, bg_h))

                try:
                    self.label_fundo.configure(image=self.fundo_ctkimage)
                except Exception:
                    pass
                try:
                    self.label_fundo.place(relx=0.5, rely=0.5, anchor='center')
                except Exception:
                    pass
            except Exception:
                pass
        if hasattr(self, 'nome_app_img_original') and self.nome_app_img_original and getattr(self, 'label_nome_app', None):
            try:
                new_w = int(event.width * 0.7)
                new_w = min(400, max(150, new_w))
                orig_w, orig_h = self.nome_app_img_original.size
                ratio = new_w / orig_w
                new_h = max(1, int(orig_h * ratio))
                logo_red = self.nome_app_img_original.resize(
                    (new_w, new_h), Image.LANCZOS)
                self.nome_app_ctkimage = ctk.CTkImage(
                    logo_red, size=(new_w, new_h))
                try:
                    self.label_nome_app.configure(image=self.nome_app_ctkimage)
                except Exception:
                    pass
                try:
                    self.label_nome_app.place(
                        relx=0.54, rely=0.185, anchor='center')
                except Exception:
                    pass
            except Exception:
                pass

    def _configurar_placeholder_manual(self, entry_widget, texto_placeholder, eh_senha=False):

        entry_widget._placeholder_text = texto_placeholder
        entry_widget._is_password_field = bool(eh_senha)
        entry_widget._placeholder_active = False

        def ativar_placeholder(event=None):
            try:
                if entry_widget.get() == "":
                    entry_widget.delete(0, 'end')
                    entry_widget.insert(0, entry_widget._placeholder_text)
                    if entry_widget._is_password_field:
                        entry_widget.configure(show='')
                    try:
                        entry_widget.configure(
                            text_color=self._PLACEHOLDER_COLOR)
                    except Exception:
                        pass
                    entry_widget._placeholder_active = True
            except Exception:
                pass

        def desativar_placeholder(event=None):
            try:
                if entry_widget._placeholder_active:
                    entry_widget.delete(0, 'end')
                    if entry_widget._is_password_field:
                        entry_widget.configure(show='*')
                    try:
                        entry_widget.configure(text_color=self._TEXT_COLOR)
                    except Exception:
                        pass
                    entry_widget._placeholder_active = False
            except Exception:
                pass

        def on_keypress(event=None):
            try:
                if getattr(entry_widget, "_placeholder_active", False):
                    entry_widget.delete(0, 'end')
                    if entry_widget._is_password_field:
                        entry_widget.configure(show='*')
                    try:
                        entry_widget.configure(text_color=self._TEXT_COLOR)
                    except Exception:
                        pass
                    entry_widget._placeholder_active = False
            except Exception:
                pass

        try:
            entry_widget.bind("<FocusIn>", desativar_placeholder, add="+")
            entry_widget.bind("<FocusOut>", ativar_placeholder, add="+")
            entry_widget.bind("<Key>", on_keypress, add="+")
        except Exception:
            pass
        ativar_placeholder()

    def criar_frames_views(self):
        # ---------- FRAME LOGIN ----------
        self.frame_login = ctk.CTkFrame(self, height=200, fg_color="#074104")
        for col in (0, 2):
            self.frame_login.grid_columnconfigure(col, weight=1)
        self.frame_login.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(self.frame_login, width=350, text='Entrar',
                     font=ctk.CTkFont('Century Gothic',
                                      size=20, weight='bold'),
                     fg_color='#FDD835', text_color='#074104').grid(row=0, column=1, pady=(0, 5))

        self.usuario_login_entry = ctk.CTkEntry(self.frame_login, width=150, height=20,
                                                font=ctk.CTkFont(
                                                    'Century Gothic', size=14),
                                                placeholder_text='', placeholder_text_color='#9E9E9E', text_color="#FFFFFF")
        self.usuario_login_entry.grid(row=1, column=1, padx=10, pady=(5, 0))

        self.senha_login_entry = ctk.CTkEntry(self.frame_login, width=150, height=20,
                                              font=ctk.CTkFont(
                                                  'Century Gothic', size=14),
                                              placeholder_text='', placeholder_text_color='#9E9E9E',
                                              corner_radius=10, show='*', text_color="#FFFFFF")
        self.senha_login_entry.grid(row=2, column=1, padx=10, pady=(5, 0))

        self._configurar_placeholder_manual(
            self.usuario_login_entry, "Usuário", eh_senha=False)
        self._configurar_placeholder_manual(
            self.senha_login_entry, "Senha", eh_senha=True)

        ctk.CTkCheckBox(self.frame_login, text=' Ver senha ',
                        font=ctk.CTkFont('Century Gothic', size=12),
                        text_color="#FFFFFF",
                        checkbox_height=15, checkbox_width=15,
                        variable=self.var_ver_senha_login,
                        command=self.toggle_show_password_login).grid(row=3, column=1, padx=(0, 0), pady=(5, 0), sticky='s')

        ctk.CTkButton(self.frame_login, text='Login', width=150, height=20,
                      font=ctk.CTkFont('Century Gothic',
                                       size=20, weight='bold'),
                      fg_color='#FDD835', hover_color='#D4B71F', text_color='#074104',
                      command=self.entrar).grid(row=4, column=1, padx=10, pady=(5, 5))

        wrapper = ctk.CTkFrame(self.frame_login, fg_color='transparent')
        wrapper.grid(row=5, column=1, pady=(0, 5))
        ctk.CTkLabel(wrapper, text='Primeiro acesso?', font=ctk.CTkFont(
            'Century Gothic', size=10), text_color="#FFFFFF").grid(row=0, column=0)
        ctk.CTkButton(wrapper, text='Cadastre-se aqui', width=100, height=15,
                      font=ctk.CTkFont('Century Gothic',
                                       size=10, weight='bold'),
                      fg_color='#2E7D32', hover_color='#063004', text_color='#FFFFFF',
                      command=self.mostrar_tela_cadastro).grid(row=0, column=1, padx=(5, 0))

        # ---------- FRAME CADASTRO ----------
        self.frame_cadastro = ctk.CTkFrame(
            self, height=410, fg_color="#074104")
        for col in (0, 2):
            self.frame_cadastro.grid_columnconfigure(col, weight=1)
        self.frame_cadastro.grid_columnconfigure(1, weight=0)

        ctk.CTkLabel(self.frame_cadastro, width=350, text='Cadastro',
                     font=ctk.CTkFont('Century Gothic',
                                      size=20, weight='bold'),
                     fg_color='#FDD835', text_color='#074104').grid(row=0, column=1, pady=(0, 5))

        self.usuario_cadastro_entry = ctk.CTkEntry(self.frame_cadastro, width=150, height=20, font=ctk.CTkFont('Century Gothic', size=14),
                                                   placeholder_text='', placeholder_text_color=self._PLACEHOLDER_COLOR, text_color="#FFFFFF")
        self.usuario_cadastro_entry.grid(row=1, column=1, padx=10, pady=(5, 0))

        self.email_cadastro_entry = ctk.CTkEntry(self.frame_cadastro, width=150, height=20, font=ctk.CTkFont('Century Gothic', size=14),
                                                 placeholder_text='', placeholder_text_color=self._PLACEHOLDER_COLOR, text_color="#FFFFFF")
        self.email_cadastro_entry.grid(row=2, column=1, padx=10, pady=(5, 0))

        self.wrapper2 = ctk.CTkFrame(
            self.frame_cadastro, fg_color='transparent')
        self.wrapper2.grid(row=3, column=1, pady=(5, 0))

        blocos = ['A', 'B', 'C', 'D']
        self.bloco_usuario_cadastro_entry = ctk.CTkComboBox(self.wrapper2, width=70, height=20, values=blocos,
                                                            font=ctk.CTkFont(
                                                                'Century Gothic', size=14),
                                                            text_color="#9E9E9E")

        self.bloco_usuario_cadastro_entry.set('Bloco')
        self.bloco_usuario_cadastro_entry.grid(
            row=0, column=0, padx=5, pady=(5, 0))

        self.apartamento_usuario_cadastro_entry = ctk.CTkEntry(self.wrapper2, width=70, height=20,
                                                               font=ctk.CTkFont(
                                                                   'Century Gothic', size=14),
                                                               placeholder_text='Apto', placeholder_text_color=self._PLACEHOLDER_COLOR,
                                                               text_color="#000000")
        self.apartamento_usuario_cadastro_entry.grid(
            row=0, column=1, padx=5, pady=(5, 0))

        self.senha_cadastro_entry = ctk.CTkEntry(self.frame_cadastro, width=150, height=20, font=ctk.CTkFont('Century Gothic', size=14),
                                                 placeholder_text='', placeholder_text_color=self._PLACEHOLDER_COLOR, corner_radius=10, show='*', text_color="#000000")
        self.senha_cadastro_entry.grid(row=4, column=1, padx=10, pady=(5, 0))

        self.confirmacao_senha_cadastro_entry = ctk.CTkEntry(self.frame_cadastro, width=150, height=20, font=ctk.CTkFont('Century Gothic',
                                                                                                                         size=14), placeholder_text='', placeholder_text_color=self._PLACEHOLDER_COLOR, corner_radius=10, show='*', text_color="#000000")
        self.confirmacao_senha_cadastro_entry.grid(
            row=5, column=1, padx=10, pady=(5, 5))

        self._configurar_placeholder_manual(
            self.usuario_cadastro_entry, "Usuário", eh_senha=False)
        self._configurar_placeholder_manual(
            self.email_cadastro_entry, "E-mail", eh_senha=False)
        self._configurar_placeholder_manual(
            self.senha_cadastro_entry, "Senha", eh_senha=True)
        self._configurar_placeholder_manual(
            self.confirmacao_senha_cadastro_entry, "Confirme a senha", eh_senha=True)

        self._configurar_placeholder_manual(
            self.apartamento_usuario_cadastro_entry, "Apto", eh_senha=False)

        ctk.CTkCheckBox(self.frame_cadastro, text=' Ver senha ',
                        font=ctk.CTkFont('Century Gothic', size=12),
                        text_color="#FFFFFF",  # checkbox text em preto para contraste sobre branco
                        checkbox_height=15, checkbox_width=15,
                        variable=self.var_ver_senha_cadastro,
                        command=self.toggle_show_password_cadastro).grid(row=6, column=1, padx=(0, 0), pady=(5, 0), sticky='s')

        ctk.CTkButton(self.frame_cadastro, text='Cadastrar', width=150, height=20,
                      font=ctk.CTkFont('Century Gothic',
                                       size=20, weight='bold'),
                      fg_color='#FDD835', hover_color='#D4B71F', text_color='#074104',
                      command=self.cadastrar).grid(row=7, column=1, padx=10, pady=(5, 5))

        wrapper3 = ctk.CTkFrame(self.frame_cadastro, fg_color='transparent')
        wrapper3.grid(row=8, column=1, pady=(0, 5))
        ctk.CTkLabel(wrapper3, text='Já tem cadastro?', font=ctk.CTkFont(
            'Century Gothic', size=10), text_color="#FFFFFF").grid(row=0, column=0)
        ctk.CTkButton(wrapper3, text='Entrar', width=75, height=15,
                      font=ctk.CTkFont('Century Gothic',
                                       size=10, weight='bold'),
                      fg_color='#2E7D32', hover_color="#063004", text_color='#FFFFFF',
                      command=self.mostrar_tela_login).grid(row=0, column=1, padx=(5, 5))

        # ---------- FRAME HOME ----------
        self.frame_home = ctk.CTkFrame(
            self, fg_color="#FFFFFF", corner_radius=0)
        try:
            self.frame_home.grid_columnconfigure(0, weight=1)
        except Exception:
            pass

        self.frame_home.grid_columnconfigure(1, weight=0)

        self.ola_label = ctk.CTkLabel(self.frame_home, width=350, height=40, text=' Olá,',
                                      font=ctk.CTkFont(
                                          'Century Gothic', size=16, weight='bold'),
                                      fg_color='#FDD835', text_color='#074104', anchor='w', corner_radius=10)
        self.ola_label.grid(row=0, column=0, pady=(10, 0),
                            padx=(10, 0), sticky='ew')

        self.sair_btn_home = ctk.CTkButton(self.frame_home, text='Sair', width=65, height=40,
                                           font=ctk.CTkFont(
                                               'Century Gothic', size=20, weight='bold'),
                                           fg_color='#074104',
                                           hover_color="#042102", text_color="#FFFFFF", corner_radius=10,
                                           command=self.mostrar_tela_login)

        self.sair_btn_home.grid(row=0, column=1, pady=(10, 0), padx=(0, 10))

        self.infos_gerais_frame = ctk.CTkFrame(
            self.frame_home, fg_color='transparent')
        self.infos_gerais_frame.grid(
            row=1, column=0, pady=(10, 0), padx=(10, 0), sticky='ew')

        self.pontuacao_label = ctk.CTkLabel(self.infos_gerais_frame, width=125, height=40,
                                            text=' Pontuação: 0', text_color='#074104', fg_color='#FDD835', corner_radius=10, anchor='w', font=ctk.CTkFont('Century Gothic', size=14, weight='bold'))
        self.pontuacao_label.grid(row=0, column=0, pady=5, padx=0, sticky='w')

        self.somatorio_container = ctk.CTkFrame(
            self.frame_home, fg_color="#074104", corner_radius=10)
        self.somatorio_container.grid(
            row=2, column=0, columnspan=2, pady=(20, 0), padx=(10, 10), sticky='ew')

        self.total_bloco_label = ctk.CTkLabel(self.somatorio_container, text="Total reciclado do seu bloco:",
                                              font=ctk.CTkFont(
                                                  'Century Gothic', size=14, weight='bold'),
                                              fg_color="#074104",
                                              text_color="#FFFFFF", anchor='w')
        self.total_bloco_label.pack(padx=10, pady=(10, 5), fill='x')

        self.somatorio_grid = ctk.CTkFrame(
            self.somatorio_container, fg_color='transparent')
        self.somatorio_grid.pack(
            pady=(5, 10), padx=10, fill='both', expand=True)

        n_rows, n_cols = 2, 3
        self.tipos_lixo_list = ['Plástico', 'Papel',
                                'Vidro', 'Lata', 'Metal', 'Eletronico']
        self.somatorio_labels = {}

        for col in range(n_cols):
            self.somatorio_grid.grid_columnconfigure(col, weight=1)

        label_font = ctk.CTkFont('Century Gothic', size=12, weight='bold')

        for r in range(n_rows):
            for c in range(n_cols):
                idx = r * n_cols + c
                if idx < len(self.tipos_lixo_list):
                    tipo = self.tipos_lixo_list[idx]
                    lbl = ctk.CTkLabel(self.somatorio_grid, text=f"{tipo}:\n0.0 kg", width=120, height=50,
                                       fg_color='#FDD835', text_color='#074104', corner_radius=10, anchor='center', font=label_font)
                    lbl.grid(row=r, column=c, padx=5, pady=5, sticky='ew')
                    self.somatorio_labels[tipo] = lbl

        self.form_descarte_frame = ctk.CTkFrame(
            self.frame_home, fg_color='#074104', height=400, corner_radius=10)
        self.form_descarte_frame.grid(
            row=3, column=0, columnspan=2, pady=15, padx=(10, 10), sticky="ew")
        self.form_descarte_frame.grid_columnconfigure(0, weight=1)

        self.descarte_frame = ctk.CTkFrame(
            self.form_descarte_frame, fg_color='transparent', corner_radius=10)
        self.descarte_frame.grid(row=0, column=0, pady=(
            8, 6), padx=(12, 12), sticky='ew')
        self.descarte_frame.grid_columnconfigure(0, weight=1)
        self.descarte_frame.grid_columnconfigure(1, weight=1)

        self.descarte_btn_home = ctk.CTkButton(self.descarte_frame, text='Descarte', width=170, height=50,
                                               font=ctk.CTkFont(
                                                   'Century Gothic', size=14, weight='bold'),
                                               fg_color='#FDD835', hover_color='#D4B71F', text_color='#074104', corner_radius=10,
                                               command=self.mostrar_tela_descarte)
        self.descarte_btn_home.grid(
            row=0, column=0, pady=(0, 0), padx=6, sticky='ew')

        self.resgate_btn_home = ctk.CTkButton(self.descarte_frame, text='Resgate', width=170, height=50,
                                              font=ctk.CTkFont(
                                                  'Century Gothic', size=14, weight='bold'),
                                              fg_color='#FDD835', hover_color='#D4B71F', text_color='#074104', corner_radius=10,
                                              command=self.mostrar_tela_resgate)
        self.resgate_btn_home.grid(
            row=0, column=1, pady=(0, 0), padx=6, sticky='ew')

        self.form_inner = ctk.CTkFrame(
            self.form_descarte_frame, fg_color='transparent')
        self.form_inner.grid(row=1, column=0, pady=(
            8, 10), padx=(20, 20), sticky='ew')
        self.form_inner.grid_columnconfigure(0, weight=1)

        tipos_lixo = ['Plástico', 'Papel',
                      'Vidro', 'Lata', 'Metal', 'Eletronico']
        self.tipo_lixo = ctk.CTkComboBox(
            self.form_inner, height=30, values=tipos_lixo, bg_color="#074104", text_color='#9E9E9E')
        self.tipo_lixo.set('Tipo de lixo')
        self.tipo_lixo.grid(row=0, column=0, pady=5, padx=(0, 0), sticky='ew')

        self.peso_lixo_entry = ctk.CTkEntry(
            self.form_inner, height=30, bg_color="#074104", text_color='#9E9E9E', placeholder_text='')

        try:
            self._configurar_placeholder_manual(
                self.peso_lixo_entry, 'Peso (kg)', eh_senha=False)
        except Exception:
            pass
        self.peso_lixo_entry.grid(
            row=1, column=0, pady=5, padx=(0, 0), sticky='ew')

        self.botao_descarte = ctk.CTkButton(self.form_inner, text='Descartar', height=34,
                                            fg_color='#FDD835', hover_color='#D4B71F', text_color='#074104', corner_radius=10,
                                            command=self.handle_descarte)
        self.botao_descarte.grid(row=2, column=0, pady=(
            15, 5), padx=(0, 0), sticky='ew')

        # ---------- FRAME RESGATE ----------
        self.frame_resgate = ctk.CTkFrame(
            self.form_descarte_frame, fg_color='transparent',)

        self.resgate_title = ctk.CTkLabel(self.frame_resgate, text="Recompensas disponíveis",
                                          font=ctk.CTkFont(
                                              'Century Gothic', size=14, weight='bold'),
                                          fg_color='transparent', text_color='#FFFFFF')
        self.resgate_title.pack(pady=(8, 6))

        self.resgate_table = ctk.CTkFrame(
            self.frame_resgate, fg_color='transparent')
        self.resgate_table.pack(padx=0, pady=(
            4, 10), fill='both', expand=True)

        for i, item in enumerate(self.recompensas):
            row_frame = ctk.CTkFrame(
                self.resgate_table, fg_color='transparent')

            row_frame.grid_columnconfigure(0, weight=1)
            row_frame.grid_columnconfigure(1, weight=0)
            row_frame.grid_columnconfigure(2, weight=0)
            row_frame.grid(row=i, column=0, sticky='ew', pady=6,
                           padx=0)

            nome_text = str(item.get('nome', '')).strip()

            nome_lbl = ctk.CTkLabel(
                row_frame,
                text=nome_text,
                anchor='w',
                font=ctk.CTkFont('Century Gothic', size=12),
                fg_color='transparent',
                text_color='#FFFFFF',
                wraplength=150,
                justify='left'
            )
            nome_lbl.grid(row=0, column=0, sticky='w', padx=(0, 8))

            custo_lbl = ctk.CTkLabel(
                row_frame,
                text=f"{int(item.get('custo', 0))} pts",
                anchor='center',
                width=60,
                font=ctk.CTkFont('Century Gothic', size=10, weight='bold'),
                fg_color='#FDD835',
                text_color='#074104',
                corner_radius=8
            )
            custo_lbl.grid(row=0, column=1, sticky='e', padx=(0, 8))

            btn = ctk.CTkButton(
                row_frame,
                text='Resgatar',
                width=60,
                height=28,
                corner_radius=8,
                command=lambda it=item: self.tentar_resgatar(
                    it['nome'], it['custo'])
            )
            btn.grid(row=0, column=2, sticky='e')

    # ---------- Troca de telas----------

    def mostrar_tela_login(self):
        if self.frame_cadastro and self.frame_cadastro.winfo_ismapped():
            self.frame_cadastro.grid_forget()
            self.limpar_dados_cadastro()
        if self.frame_home and self.frame_home.winfo_ismapped():
            self.frame_home.grid_forget()

        self.configure(fg_color=self._default_bg)

        try:
            if getattr(self, 'label_fundo', None):
                try:
                    self.label_fundo.place(relx=0.5, rely=0.5, anchor='center')
                except Exception:
                    try:
                        if getattr(self, 'fundo_img_original', None):
                            cur_w = self.winfo_width() if self.winfo_width() > 1 else 350
                            cur_h = self.winfo_height() if self.winfo_height() > 1 else 600
                            target_w = max(1, int(cur_w * self._bg_scale_x))
                            target_h = max(1, int(cur_h * self._bg_scale_y))
                            bg_red = self.fundo_img_original.resize(
                                (target_w, target_h), Image.LANCZOS)
                            self.fundo_ctkimage = ctk.CTkImage(
                                bg_red, size=(target_w, target_h))
                            self.label_fundo = ctk.CTkLabel(
                                self, text='', image=self.fundo_ctkimage, fg_color='transparent')
                            self.label_fundo.place(
                                relx=0.5, rely=0.5, anchor='center')
                    except Exception:
                        pass

            if getattr(self, 'label_nome_app', None):
                try:
                    self.label_nome_app.place(
                        relx=0.54, rely=0.185, anchor='center')
                except Exception:
                    try:
                        if getattr(self, 'nome_app_img_original', None):

                            new_w = int(self.winfo_width() *
                                        0.7) if self.winfo_width() > 1 else 250
                            new_w = min(400, max(150, new_w))
                            orig_w, orig_h = self.nome_app_img_original.size
                            ratio = new_w / orig_w
                            new_h = max(1, int(orig_h * ratio))
                            logo_red = self.nome_app_img_original.resize(
                                (new_w, new_h), Image.LANCZOS)
                            self.nome_app_ctkimage = ctk.CTkImage(
                                logo_red, size=(new_w, new_h))
                            self.label_nome_app = ctk.CTkLabel(
                                self, text='', image=self.nome_app_ctkimage, fg_color='transparent')
                            self.label_nome_app.place(
                                relx=0.54, rely=0.185, anchor='center')
                    except Exception:
                        pass
        except Exception:
            pass

        if self.frame_login:
            self.frame_login.grid(row=0, column=0, sticky='s', padx=0, pady=0)
            self.after(40, self._remover_foco_entradas)

    def mostrar_tela_cadastro(self):
        if self.frame_login and self.frame_login.winfo_ismapped():
            self.frame_login.grid_forget()
            self.limpar_dados_login()
        self.configure(fg_color="white")
        if self.frame_cadastro:
            self.frame_cadastro.grid(
                row=0, column=0, sticky='s', padx=0, pady=0)
            self.after(40, self._remover_foco_entradas)

    def mostrar_tela_home(self, user_dict: dict = None):
        try:
            if self.frame_login and self.frame_login.winfo_ismapped():
                self.frame_login.grid_forget()
            if self.frame_cadastro and self.frame_cadastro.winfo_ismapped():
                self.frame_cadastro.grid_forget()

            if hasattr(self, 'label_fundo') and self.label_fundo:
                try:
                    self.label_fundo.place_forget()

                except Exception:
                    pass
            if hasattr(self, 'label_nome_app') and self.label_nome_app:
                try:
                    self.label_nome_app.place_forget()
                except Exception:
                    pass

            if isinstance(user_dict, dict):
                self.current_user = user_dict

            if getattr(self, "current_user", None):
                nome = str(self.current_user.get('usuario')
                           or '-').strip().capitalize()
                bloco = str(self.current_user.get('bloco') or '-')
                apto = str(self.current_user.get('apartamento') or '-')
                texto = f" Olá, {nome} | Bloco: {bloco}  Apto: {apto}"
            else:
                texto = " Olá,"

            self.ola_label.configure(text=texto)

            if self.frame_home:
                self.frame_home.grid(
                    row=0, column=0, sticky='nsew', padx=0, pady=0)
                self.after(40, self._remover_foco_entradas)

            bloco_atual = self.current_user.get('bloco') if getattr(
                self, "current_user", None) else None
            if bloco_atual:
                self.atualizar_somatorio_bloco(bloco_atual)
            else:

                for tipo, lbl in self.somatorio_labels.items():
                    try:
                        lbl.configure(text=f" {tipo}:\n0.0 kg ")
                    except Exception:
                        pass

            try:
                if getattr(self, 'current_user', None) and isinstance(self.current_user, dict) and self.current_user.get('pontuacao') is not None:
                    self.pontuacao_atual = float(
                        self.current_user.get('pontuacao') or 0)
                if not hasattr(self, 'pontuacao_atual') or self.pontuacao_atual is None:
                    self.pontuacao_atual = 0
                if float(self.pontuacao_atual).is_integer():
                    self.pontuacao_label.configure(
                        text=f' Pontuação: {int(self.pontuacao_atual)}')
                else:
                    self.pontuacao_label.configure(
                        text=f' Pontuação: {self.pontuacao_atual:.1f}')
            except Exception:
                pass

            try:
                if getattr(self, 'frame_resgate', None) and self.frame_resgate.winfo_ismapped():
                    self.frame_resgate.pack_forget() if hasattr(
                        self.frame_resgate, 'pack_forget') else None
                    try:
                        self.frame_resgate.grid_forget()
                    except Exception:
                        pass
                if getattr(self, 'form_inner', None) and not self.form_inner.winfo_ismapped():

                    try:
                        self.form_inner.grid(row=1, column=0, pady=(
                            8, 10), padx=(20, 20), sticky='ew')
                    except Exception:
                        pass
            except Exception:
                pass

        except Exception as e:
            print("Erro ao mostrar Home:", e)

    def mostrar_tela_descarte(self):

        if getattr(self, 'somatorio_container', None) and not self.somatorio_container.winfo_ismapped():
            self.somatorio_container.grid(
                row=2, column=0, columnspan=2, pady=(20, 0), padx=(10, 10), sticky='ew')

        try:

            try:
                if getattr(self, 'frame_resgate', None) and self.frame_resgate.winfo_ismapped():

                    try:
                        self.frame_resgate.pack_forget()
                    except Exception:
                        pass
                    try:
                        self.frame_resgate.grid_forget()
                    except Exception:
                        pass
            except Exception:
                pass

            try:
                if getattr(self, 'form_inner', None) and not self.form_inner.winfo_ismapped():
                    self.form_inner.grid(row=1, column=0, pady=(
                        8, 10), padx=(20, 20), sticky='ew')
            except Exception:
                pass
        except Exception as e:
            print("Erro ao mostrar tela de descarte:", e)

    def mostrar_tela_resgate(self, user_dict: dict = None):

        try:
            if self.frame_login and self.frame_login.winfo_ismapped():
                self.frame_login.grid_forget()
            if self.frame_cadastro and self.frame_cadastro.winfo_ismapped():
                self.frame_cadastro.grid_forget()

            try:
                if getattr(self, 'form_inner', None) and self.form_inner.winfo_ismapped():
                    try:
                        self.form_inner.grid_forget()
                    except Exception:
                        pass
            except Exception:
                pass

            if isinstance(user_dict, dict):
                self.current_user = user_dict

            if getattr(self, 'somatorio_container', None) and self.somatorio_container.winfo_ismapped():
                self.somatorio_container.grid_forget()

            try:
                if getattr(self, 'frame_resgate', None) and not self.frame_resgate.winfo_ismapped():

                    self.frame_resgate.grid(row=1, column=0, pady=(
                        8, 10), padx=(20, 20), sticky='nsew')
            except Exception:
                pass

            try:
                if getattr(self, "current_user", None):
                    nome = str(self.current_user.get('usuario')
                               or '-').strip().capitalize()
                    bloco = str(self.current_user.get('bloco') or '-')
                    apto = str(self.current_user.get('apartamento') or '-')
                    texto = f" Olá, {nome} | Bloco: {bloco}    Apto: {apto}"
                else:
                    texto = " Olá,"
                self.ola_label.configure(text=texto)
            except Exception:
                pass

        except Exception as e:
            print("Erro ao mostrar Resgate:", e)

    def mostrar_alerta(self, mensagem, cor_fundo="#FDD835", duracao=2000, frame=None):

        try:
            target = frame if frame else self

            alerta_bg = ctk.CTkFrame(
                target, fg_color=cor_fundo, bg_color=cor_fundo, corner_radius=10, border_width=0)

            alerta_bg.place(relx=0.5, y=15, anchor="n")

            alerta_lbl = ctk.CTkLabel(
                alerta_bg,
                text=mensagem,
                fg_color="transparent",
                text_color="#FFFFFF",
                font=ctk.CTkFont("Century Gothic", 11,
                                 "bold"),
                anchor="center"
            )
            alerta_lbl.pack(padx=8, pady=5)

            try:
                alerta_bg.lift()
            except Exception:
                pass

            def _destroy_alert():
                try:
                    alerta_bg.destroy()
                except Exception:
                    pass

            self.after(duracao, _destroy_alert)
        except Exception as e:
            print("Erro ao mostrar alerta:", e)

    def _forcar_reativacao_placeholder(self, entry_widget):
        if entry_widget.get() == "":
            entry_widget.delete(0, 'end')
            entry_widget.insert(0, entry_widget._placeholder_text)
            if getattr(entry_widget, '_is_password_field', False):
                entry_widget.configure(show='')
            try:
                entry_widget.configure(text_color=self._PLACEHOLDER_COLOR)
            except Exception:
                pass
            entry_widget._placeholder_active = True

    def limpar_dados_cadastro(self):
        try:
            self.usuario_cadastro_entry.delete(0, END)
            self.email_cadastro_entry.delete(0, END)
            self.senha_cadastro_entry.delete(0, END)
            self.confirmacao_senha_cadastro_entry.delete(0, END)

            try:
                self.bloco_usuario_cadastro_entry.set('Bloco')
            except Exception:
                pass
            self.apartamento_usuario_cadastro_entry.delete(0, END)

            self.var_ver_senha_cadastro.set(False)
            self.toggle_show_password_cadastro()

            for entry in [
                self.usuario_cadastro_entry,
                self.email_cadastro_entry,
                self.senha_cadastro_entry,
                self.confirmacao_senha_cadastro_entry,
                self.apartamento_usuario_cadastro_entry
            ]:
                self._forcar_reativacao_placeholder(entry)

        except Exception as e:
            print("Erro ao limpar cadastro:", e)

    def limpar_dados_login(self):
        try:
            self.usuario_login_entry.delete(0, END)
            self.senha_login_entry.delete(0, END)
            self.var_ver_senha_login.set(False)
            self.toggle_show_password_login()
            for entry in [self.usuario_login_entry, self.senha_login_entry]:
                self._forcar_reativacao_placeholder(entry)
        except Exception as e:
            print("Erro ao limpar login:", e)

    def _remover_foco_entradas(self):

        try:
            self.focus_set()
        except Exception:

            pass

    def toggle_show_password_login(self):
        try:
            if self.var_ver_senha_login.get():
                self.senha_login_entry.configure(show='')
            else:
                if getattr(self.senha_login_entry, "_placeholder_active", False):
                    self.senha_login_entry.configure(show='')
                else:
                    self.senha_login_entry.configure(show='*')
        except Exception as e:
            print("Erro toggle login:", e)

    def toggle_show_password_cadastro(self):
        try:
            if self.var_ver_senha_cadastro.get():
                self.senha_cadastro_entry.configure(show='')
                self.confirmacao_senha_cadastro_entry.configure(show='')
            else:
                if getattr(self.senha_cadastro_entry, "_placeholder_active", False):
                    self.senha_cadastro_entry.configure(show='')
                else:
                    self.senha_cadastro_entry.configure(show='*')
                if getattr(self.confirmacao_senha_cadastro_entry, "_placeholder_active", False):
                    self.confirmacao_senha_cadastro_entry.configure(show='')
                else:
                    self.confirmacao_senha_cadastro_entry.configure(show='*')
        except Exception as e:
            print("Erro toggle cadastro:", e)

    # ---------- Handlers ----------
    def cadastrar(self):
        usuario = self.usuario_cadastro_entry.get().strip()
        bloco = self.bloco_usuario_cadastro_entry.get().strip() if hasattr(
            self, 'bloco_usuario_cadastro_entry') else ''
        apartamento_raw = self.apartamento_usuario_cadastro_entry.get().strip(
        ) if hasattr(self, 'apartamento_usuario_cadastro_entry') else ''
        email = self.email_cadastro_entry.get().strip()
        senha = self.senha_cadastro_entry.get()
        confirmar = self.confirmacao_senha_cadastro_entry.get()

        if getattr(self.usuario_cadastro_entry, "_placeholder_active", False) and usuario == self.usuario_cadastro_entry._placeholder_text:
            usuario = ""
        if getattr(self.email_cadastro_entry, "_placeholder_active", False) and email == self.email_cadastro_entry._placeholder_text:
            email = ""
        if getattr(self.senha_cadastro_entry, "_placeholder_active", False) and senha == self.senha_cadastro_entry._placeholder_text:
            senha = ""
        if getattr(self.confirmacao_senha_cadastro_entry, "_placeholder_active", False) and confirmar == self.confirmacao_senha_cadastro_entry._placeholder_text:
            confirmar = ""

        if bloco == 'Bloco':
            bloco = ""
        if getattr(self.apartamento_usuario_cadastro_entry, "_placeholder_active", False) and apartamento_raw == self.apartamento_usuario_cadastro_entry._placeholder_text:
            apartamento_raw = ""

        if not usuario or not email or not senha:
            self.mostrar_alerta(
                'Preencha todos os campos obrigatórios', cor_fundo="#CE0606")
            return

        if senha != confirmar:
            self.mostrar_alerta(
                'Senhas diferentes, tente novamente.', cor_fundo="#CE0606")
            return

        apartamento = None
        if apartamento_raw:
            try:
                apartamento = int(apartamento_raw)
            except ValueError:
                self.mostrar_alerta(
                    "Apto deve ser um número inteiro. Use apenas dígitos.", cor_fundo="#CE0606")
                return

        try:
            df = self.ler_usuarios_df()
            if not df.empty:
                if ((df['usuario'].astype(str).str.lower() == str(usuario).lower()).any() and usuario):
                    self.mostrar_alerta(
                        "Nome de usuário já cadastrado. Tente outro.", cor_fundo="#CE0606")
                    return
                if ((df['email'].astype(str).str.lower() == str(email).lower()).any() and email):
                    self.mostrar_alerta(
                        "E-mail já cadastrado. Tente outro.", cor_fundo="#CE0606")
                    return
        except Exception as e:
            print("Erro ao verificar duplicidade:", e)

        try:
            self.inserir_usuario_excel(
                usuario, bloco if bloco else None, apartamento, email, senha)
            self.mostrar_alerta(
                'Cadastro realizado com sucesso.', cor_fundo="#3FCE06")
            print('Dados de usuário inseridos com sucesso no Excel!')
        except Exception as e:
            print('Erro ao inserir no Excel:', e)
            self.mostrar_alerta('Erro ao salvar usuário.', cor_fundo="#CE0606")
            return

        self.limpar_dados_cadastro()
        self.mostrar_tela_login()

    def entrar(self):
        usuario_ou_email = self.usuario_login_entry.get().strip()
        senha = self.senha_login_entry.get()

        if getattr(self.usuario_login_entry, "_placeholder_active", False) and usuario_ou_email == self.usuario_login_entry._placeholder_text:
            usuario_ou_email = ""
        if getattr(self.senha_login_entry, "_placeholder_active", False) and senha == self.senha_login_entry._placeholder_text:
            senha = ""

        if not usuario_ou_email or not senha:
            self.mostrar_alerta(
                "Preencha com seu usuário e senha.", cor_fundo="#CE0606")
            return

        try:
            df = self.ler_usuarios_df()
            if df.empty:
                self.mostrar_alerta(
                    "Nenhum usuário cadastrado. Cadastre-se primeiro.", cor_fundo="#CE0606")
                return

            df['usuario_str'] = df['usuario'].astype(
                str).fillna('').str.strip()
            df['email_str'] = df['email'].astype(str).fillna('').str.strip()
            df['senha_str'] = df['senha'].astype(str).fillna('')

            busc = str(usuario_ou_email).strip()

            match_idx = df.index[
                (df['usuario_str'].str.lower() == busc.lower()) |
                (df['email_str'].str.lower() == busc.lower())
            ].tolist()

            if not match_idx:
                self.mostrar_alerta(
                    "Usuário não encontrado.", cor_fundo="#CE0606")
                return

            idx = match_idx[0]
            senha_armazenada = df.at[idx, 'senha_str']

            if str(senha) == str(senha_armazenada):
                self.current_user = df.loc[idx].to_dict()

                try:
                    valor = df.at[idx,
                                  'pontuacao'] if 'pontuacao' in df.columns else 0
                    self.pontuacao_atual = float(
                        valor) if not pd.isna(valor) else 0
                except Exception:
                    self.pontuacao_atual = getattr(self, 'pontuacao_atual', 0)

                try:
                    if float(self.pontuacao_atual).is_integer():
                        self.pontuacao_label.configure(
                            text=f' Pontuação: {int(self.pontuacao_atual)}')
                    else:
                        self.pontuacao_label.configure(
                            text=f' Pontuação: {self.pontuacao_atual:.1f}')
                except Exception:
                    pass

                self.mostrar_alerta("Login efetuado.", cor_fundo="#3FCE06")
                self.limpar_dados_login()
                self.mostrar_tela_home()
            else:
                self.mostrar_alerta("Senha incorreta.", cor_fundo="#CE0606")
                return

        except Exception as e:
            print("Erro no processo de login:", e)
            return

    def handle_descarte(self):

        tipo = ''
        try:
            tipo = self.tipo_lixo.get() if hasattr(self, 'tipo_lixo') else ''
        except Exception:
            tipo = ''

        peso_raw = ''
        try:
            peso_raw = self.peso_lixo_entry.get().strip() if hasattr(self,
                                                                     'peso_lixo_entry') else ''
        except Exception:
            peso_raw = ''

        if not tipo or tipo.strip() == '' or tipo == 'Tipo de lixo':
            self.mostrar_alerta(
                "Selecione um tipo de lixo.", cor_fundo="#CE0606")
            return

        if getattr(self.peso_lixo_entry, "_placeholder_active", False) and peso_raw == getattr(self.peso_lixo_entry, "_placeholder_text", ''):
            peso_raw = ''

        try:
            peso = float(peso_raw.replace(',', '.'))
            if peso <= 0:
                self.mostrar_alerta(
                    "Informe um peso maior que zero.", cor_fundo="#CE0606",)
                return
        except Exception:
            self.mostrar_alerta(
                "Peso inválido. Use números (ex: 1.5).", cor_fundo="#CE0606")
            return

        pontos = self.pontos_por_tipo.get(tipo, 0) * peso
        try:

            if not hasattr(self, 'pontuacao_atual') or self.pontuacao_atual is None:
                self.pontuacao_atual = 0
            self.pontuacao_atual += pontos
        except Exception:

            self.pontuacao_atual = pontos

        try:

            if float(self.pontuacao_atual).is_integer():
                self.pontuacao_label.configure(
                    text=f' Pontuação: {int(self.pontuacao_atual)}')
            else:
                self.pontuacao_label.configure(
                    text=f' Pontuação: {self.pontuacao_atual:.1f}')
        except Exception:
            pass

        if not getattr(self, 'current_user', None):
            self.mostrar_alerta("Nenhum usuário logado.", cor_fundo="#CE0606")
            return

        usuario = self.current_user.get('usuario') or ''
        bloco = self.current_user.get('bloco') or ''
        apto = self.current_user.get('apartamento') or ''

        ok = self.gravar_descarte(
            tipo_lixo=tipo, peso_kg=peso, usuario=usuario, bloco=bloco, apartamento=apto)
        if ok:
            print(f"Descarte registrado: {tipo} — {peso} kg (Bloco {bloco})")

            try:
                self.peso_lixo_entry.delete(0, END)

                self._forcar_reativacao_placeholder(self.peso_lixo_entry)
            except Exception:
                pass

            try:
                self.tipo_lixo.set('Tipo de lixo')
            except Exception:
                pass

            if bloco:
                self.atualizar_somatorio_bloco(bloco)

            try:
                if isinstance(self.current_user, dict):
                    self.current_user['pontuacao'] = self.pontuacao_atual
            except Exception:
                pass

            try:

                uid = self.current_user.get('id') if isinstance(
                    self.current_user, dict) else None
                if uid:
                    self.salvar_pontuacao_usuario(uid)
                else:
                    self.salvar_pontuacao_usuario()
            except Exception:
                pass

    def tentar_resgatar(self, nome_recompensa: str, custo: float):

        try:
            if not getattr(self, 'current_user', None):
                self.mostrar_alerta(
                    "Faça login para resgatar recompensas.", cor_fundo="#CE0606")
                return

            if not hasattr(self, 'pontuacao_atual') or self.pontuacao_atual is None:
                self.pontuacao_atual = 0.0

            if float(self.pontuacao_atual) < float(custo):
                self.mostrar_alerta(
                    f'Pontos insuficientes para:\n "{nome_recompensa}".\n ({int(custo)} pts)', cor_fundo="#CE0606")
                return

            try:
                self.pontuacao_atual = float(
                    self.pontuacao_atual) - float(custo)
                if self.pontuacao_atual < 0:
                    self.pontuacao_atual = 0
            except Exception:

                self.pontuacao_atual = 0

            try:
                if isinstance(self.current_user, dict):
                    self.current_user['pontuacao'] = self.pontuacao_atual
            except Exception:
                pass

            try:
                if float(self.pontuacao_atual).is_integer():
                    self.pontuacao_label.configure(
                        text=f' Pontuação: {int(self.pontuacao_atual)}')
                else:
                    self.pontuacao_label.configure(
                        text=f' Pontuação: {self.pontuacao_atual:.1f}')
            except Exception:
                pass

            try:
                uid = self.current_user.get('id') if isinstance(
                    self.current_user, dict) else None
                if uid:
                    self.salvar_pontuacao_usuario(uid)
                else:
                    self.salvar_pontuacao_usuario()
            except Exception:
                pass

            self.mostrar_alerta(
                f"Resgate realizado: {nome_recompensa} (-{int(custo)} pts)", cor_fundo="#3FCE06")

        except Exception as e:
            print("Erro ao tentar resgatar:", e)
            self.mostrar_alerta(
                "Erro ao processar resgate.", cor_fundo="#CE0606")

    def atualizar_somatorio_bloco(self, bloco):

        try:
            soma_dict = self.somatorio_por_tipo_por_bloco(bloco)
            for tipo in self.tipos_lixo_list:
                peso = soma_dict.get(tipo, 0.0)
                lbl = self.somatorio_labels.get(tipo)
                if lbl:
                    lbl.configure(text=f"{tipo}:\n{peso:.1f} kg")
        except Exception as e:
            print("Erro ao atualizar total na UI:", e)

            for tipo, lbl in self.somatorio_labels.items():
                lbl.configure(text=f"{tipo}:\n0.0 kg")


if __name__ == '__main__':
    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme('green')
    app = EcoHiveApp()
    app.mainloop()
