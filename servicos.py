from kivy.properties import StringProperty
from kivymd.app import MDApp
from kivymd.uix.button import MDFlatButton, MDRaisedButton
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
import os
from datetime import datetime, date
import pyodbc
from kivy.utils import get_color_from_hex
from kivymd.uix.dialog import MDDialog
import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl.reader.excel import load_workbook
from fpdf import FPDF


class ContentNavigationDrawer(Screen):
    pass


class Principal(Screen):
    descr_serv = StringProperty('')

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dialog_atu = None
        self.dialog_add = None
        self.dialog_not = None
        self.dialog_err = None
        self.dialog_apg = None
        self.dialog_obs = None
        self.cnx = None
        with open('dados.txt', 'r', encoding='utf-8') as bd:  # Caminho da pasta no servidor com o banco de dados
            dados = bd.readlines()
            self.diretorio = dados[0]
            self.diretorio = self.diretorio.rstrip().split('\\')
            self.responsavel = dados[1].split('; ')
            self.dialog = None

    def mascara(self):  # Formatar CNPJ com pontos e barra
        mask = self.ids.num_cnpj.text
        if mask != '' and '/' not in mask and len(mask) >= 14:
            mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
            self.ids.num_cnpj.text = mask_cnpj
        else:
            pass

    def busca_cadastro(self):  # Buscar os dados com o CNPJ fornecido de Nome, Situação Tributária
        if self.ids.num_cnpj.text != '' and 'aluguel' not in self.ids.num_cnpj.text.lower():  # Aluguel é Pessoa Física
            try:
                lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
                self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = self.cnx.cursor()
                cursor.execute('select nome from cadastro where cnpj = ?', (self.ids.num_cnpj.text,))
                busca_nome = cursor.fetchone()
                self.ids.cod_fornec.text = busca_nome[0]

                # Buscar Situação Tributária
                cursor.execute('select optante_simples from cadastro where cnpj = ?', (self.ids.num_cnpj.text,))
                busca_simples = cursor.fetchone()
                self.ids.regime_trib.text = busca_simples[0]
            except TypeError:
                if not self.dialog:
                    self.dialog = MDDialog(text="Fornecedor não cadastrado. Deseja cadastrar?",
                                           buttons=[MDFlatButton(text="NÃO",
                                                                 theme_text_color="Custom",
                                                                 on_press=self.fecha_dialog),
                                                    MDRaisedButton(text="SIM", theme_text_color="Custom",
                                                                   on_press=self.pega_tela), ], )
                self.dialog.open()
        else:
            pass

    def pega_tela(self, inst):  # Ir à tela de cadastro para fornecedores não cadastrados
        self.manager.current = 'tela_prest'
        self.dialog.dismiss()

    def fecha_dialog(self, inst):  # Fecha caixa de diálogo caso não deseje cadastrar
        self.dialog.dismiss()

    def busca_servico(self):  # Buscar no cadastro as aliquotas segundo o código de serviço utilizado
        lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
        self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = self.cnx.cursor()
        if self.ids.cod_serv.text != '':
            self.ids.cod_serv.text = self.ids.cod_serv.text.lstrip('0')
            lista = {'irrf': self.ids.aliq_ir, 'crf': self.ids.aliq_crf, 'inss': self.ids.aliq_inss,
                     'iss': self.ids.aliq_iss}
            for imp, aliq in lista.items():
                if self.ids.regime_trib.text in 'nãoNÃOnaoNAONãoNormalnormal':  # Caso não seja Simples Nacional
                    cursor.execute(f'select {imp} from tabela_iss where servico = ?', (self.ids.cod_serv.text,))
                    busca = cursor.fetchone()
                    aliq.text = str(round(busca[0], 2)).replace('.', ',')

                else:
                    if imp == 'iss' and self.ids.mun_iss.text != '':  # Buscar alíquota do Simples do prestador
                        try:
                            cursor.execute(f'select ALIQUOTA from cadastro where CNPJ = ?', (self.ids.num_cnpj.text,))
                            busca = cursor.fetchone()
                            aliq.text = str(round(busca[0], 2)).replace('.', ',')
                        except TypeError:
                            aliq.text = '0'
                    else:
                        aliq.text = '0'

        if self.ids.regime_trib.text not in 'Simplessimples':  # Não sendo simples, buscar aliquota da prefeitura do cad
            try:
                cursor.execute('select aliq_iss from municipios where municipio = ? and cod_iss = ?',
                               (self.ids.mun_iss.text.capitalize(), self.ids.cod_serv.text,))
                busca_aliq = cursor.fetchone()
                self.ids.aliq_iss.text = str(round(busca_aliq[0], 2)).replace('.', ',')
            except TypeError:
                pass
        try:
            cursor.execute(f'select descricao from tabela_iss where servico = ?', (self.ids.cod_serv.text,))
            busca2 = cursor.fetchone()
            self.descr_serv = busca2[0][0:190]
        except TypeError:
            pass

    def aliq_desoneracao(self):  # Empresa com desoneração a aliquota de INSS é 3,5%
        if self.ids.inss_reduzido.active is True:
            self.ids.aliq_inss.text = '3,5'
        else:
            self.ids.aliq_inss.text = '0'

    def calcula_imposto(self, instance, aliquota):  # calcular impostos com o valor bruto fornecido e aliquotas
        if aliquota.text != '':
            tupla = (aliquota.text.replace(',', '.'), self.ids.v_bruto.text.replace(',', '.'))
            instance.text = str(round(float(tupla[1]) * (float(tupla[0]) / 100), 2)).replace('.', ',')
        if aliquota.text == '11,00' or aliquota.text == '3,5':  # Construção civil
            if '%' in self.ids.exclusao.text:  # Dedução de materiais e equipamentos do valor tributado em %
                tupla = (aliquota.text.replace(',', '.'), self.ids.v_bruto.text.replace(',', '.'))
                instance.text = str(round(float(tupla[1]) * float(self.ids.exclusao.text.replace('%', '')) / 100 *
                                          (float(tupla[0]) / 100), 2)).replace('.', ',')
            else:  # Dedução de materiais e equipamentos do valor tributado em R$
                tupla = (aliquota.text.replace(',', '.'), self.ids.v_bruto.text.replace(',', '.'))
                instance.text = str(round((float(tupla[1]) - float(self.ids.exclusao.text.replace(',', '.'))) *
                                          (float(tupla[0]) / 100), 2)).replace('.', ',')
        aliq_ir_pf = ['7,50', '15,00', '22,50', '27,50']  # Para aluguéis PF, aliquotas vigentes do IRRF
        deducao = ['142,80', '354,80', '636,13', '869,36']  # Parcela a ser deduzida do cálculo
        if aliquota.text in aliq_ir_pf:
            tupla = (aliquota.text.replace(',', '.'), self.ids.v_bruto.text.replace(',', '.'))
            instance.text = str(round(float(tupla[1]) * (float(tupla[0]) / 100) -
                                      float(deducao[aliq_ir_pf.index(aliquota.text)].replace(',', '.')), 2)).replace(
                '.', ',')

    def valor_liq(self):  # Calcular valor líquido a pagar
        self.ids.v_liq.text = str(round(float(self.ids.v_bruto.text.replace(',', '.')) -
                                        (sum([float(self.ids.irrf.text.replace(',', '.')),
                                              float(self.ids.crf.text.replace(',', '.')),
                                              float(self.ids.inss.text.replace(',', '.')),
                                              float(self.ids.iss.text.replace(',', '.'))])), 2)).replace('.', ',')

    def data_dia(self):  # Trazer atual para o campo data da análise
        if self.ids.dt_nota.text == '':
            self.ids.dt_analise.text = date.today().strftime('%d/%m/%Y')
        else:
            pass

    def adicionar(self):  # Adicionar nota fiscal lançada
        if self.ids.num_cnpj.text == '':
            self.dialog_obs = MDDialog(
                text="Insira todas as informações!",
                radius=[20, 7, 20, 7], )

            self.dialog_obs.open()
        else:
            lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute(
                'INSERT INTO notas_fiscais (data_analise, data, data_vencimento, NF,	CNPJ, Fornecedor, cidade,'
                'simples_nacional, codigo_servico, valor_bruto, aliq_irrf, irrf,	aliq_crf, crf, aliq_inss, '
                'inss,	aliq_iss, iss, valor_liquido) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,'
                ' ?, ?, ?, ?)', (self.ids.dt_analise.text,
                                 self.ids.dt_nota.text,
                                 self.ids.dt_venc.text,
                                 self.ids.num_nota.text,
                                 self.ids.num_cnpj.text,
                                 self.ids.cod_fornec.text,
                                 self.ids.mun_iss.text.capitalize(),
                                 self.ids.regime_trib.text,
                                 self.ids.cod_serv.text,
                                 self.ids.v_bruto.text,
                                 self.ids.aliq_ir.text,
                                 self.ids.irrf.text,
                                 self.ids.aliq_crf.text,
                                 self.ids.crf.text,
                                 self.ids.aliq_inss.text,
                                 self.ids.inss.text,
                                 self.ids.aliq_iss.text,
                                 self.ids.iss.text,
                                 self.ids.v_liq.text))
            self.ids.lembrar.active = False
            self.ids.inss_reduzido.active = False
            self.limpar()
            cnx.commit()
            cnx.close()

            self.dialog_add = MDDialog(text="Registro incluido com sucesso!", radius=[20, 7, 20, 7], )
            self.dialog_add.open()

    def limpar(self):  # Limpar os campos
        entradas = [self.ids.dt_analise, self.ids.dt_nota,
                    self.ids.dt_venc,
                    self.ids.num_nota,
                    self.ids.num_cnpj,
                    self.ids.cod_fornec,
                    self.ids.mun_iss,
                    self.ids.regime_trib,
                    self.ids.cod_serv,
                    self.ids.v_bruto,
                    self.ids.aliq_ir,
                    self.ids.irrf,
                    self.ids.aliq_crf,
                    self.ids.crf,
                    self.ids.aliq_inss,
                    self.ids.inss,
                    self.ids.aliq_iss,
                    self.ids.iss,
                    self.ids.v_liq]

        for i in entradas:
            i.text = ''
        self.descr_serv = ''

    def apagar(self):  # Apagar nota do banco de dados
        lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('DELETE FROM notas_fiscais WHERE ID=?', (self.ids.cod_id.text,))
        cnx.commit()
        cnx.close()
        self.dialog_apg = MDDialog(text="Registro apagado com sucesso!", radius=[20, 7, 20, 7], )
        self.dialog_apg.open()
        self.limpar()

    def buscar(self):  # Pesquisar com número da nota
        try:
            lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('select * FROM notas_fiscais WHERE NF=?', (self.ids.num_nota.text,))
            row = cursor.fetchone()
            self.ids.cod_id.text = str(row[0])
            self.ids.dt_analise.text = row[1]
            self.ids.dt_nota.text = row[2]
            self.ids.dt_venc.text = row[3]
            self.ids.num_nota.text = str(row[4])
            self.ids.num_cnpj.text = row[5]
            self.ids.cod_fornec.text = row[6]
            self.ids.mun_iss.text = row[7]
            self.ids.regime_trib.text = row[8]
            self.ids.cod_serv.text = row[9]
            self.ids.v_bruto.text = str(round(row[10], 2)).replace('.', ',')
            self.ids.aliq_ir.text = str(round(row[11], 2)).replace('.', ',')
            self.ids.irrf.text = str(round(row[12], 2)).replace('.', ',')
            self.ids.aliq_crf.text = str(round(row[13], 2)).replace('.', ',')
            self.ids.crf.text = str(round(row[14], 2)).replace('.', ',')
            self.ids.aliq_inss.text = str(round(row[15], 2)).replace('.', ',')
            self.ids.inss.text = str(round(row[16], 2)).replace('.', ',')
            self.ids.aliq_iss.text = str(round(row[17], 2)).replace('.', ',')
            self.ids.iss.text = str(round(row[18], 2)).replace('.', ',')
            self.ids.v_liq.text = str(round(row[19], 2)).replace('.', ',')
            cnx.commit()
            cnx.close()
        except TypeError:
            self.dialog_not = MDDialog(text="Registro não encontrado!", radius=[20, 7, 20, 7], )
            self.dialog_not.open()
            self.limpar()

    def atualizar(self):  # Atualizar dados da nota fiscal no banco de dados
        try:
            lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute(
                'update notas_fiscais set DATA_ANALISE=?, DATA=?, DATA_VENCIMENTO=?, NF=?, CNPJ=?, FORNECEDOR=?, '
                'CIDADE=?, SIMPLES_NACIONAL=?, CODIGO_SERVICO=?, VALOR_BRUTO=?, ALIQ_IRRF=?, IRRF=?, ALIQ_CRF=?, '
                'crf=?, ALIQ_INSS=?, INSS=?, ALIQ_ISS=?, ISS=?, VALOR_LIQUIDO=? where ID=?', (self.ids.dt_analise.text,
                                                                                              self.ids.dt_nota.text,
                                                                                              self.ids.dt_venc.text,
                                                                                              self.ids.num_nota.text,
                                                                                              self.ids.num_cnpj.text,
                                                                                              self.ids.cod_fornec.text,
                                                                                              self.ids.mun_iss.text,
                                                                                              self.ids.regime_trib.text,
                                                                                              self.ids.cod_serv.text,
                                                                                              self.ids.v_bruto.text,
                                                                                              self.ids.aliq_ir.text,
                                                                                              self.ids.irrf.text,
                                                                                              self.ids.aliq_crf.text,
                                                                                              self.ids.crf.text,
                                                                                              self.ids.aliq_inss.text,
                                                                                              self.ids.inss.text,
                                                                                              self.ids.aliq_iss.text,
                                                                                              self.ids.iss.text,
                                                                                              self.ids.v_liq.text,
                                                                                              self.ids.cod_id.text))
            cnx.commit()
            cnx.close()
            self.dialog_atu = MDDialog(text="Registro alterado com sucesso!", radius=[20, 7, 20, 7], )
            self.dialog_atu.open()
            self.limpar()
            self.inserir_notas()

        except TypeError:
            self.dialog_err = MDDialog(text="Erro!", radius=[20, 7, 20, 7], )
            self.dialog_err.open()

    def inserir_notas(self):  # Inserir dados da nota a ser modificada
        entradas = [self.ids.cod_id, self.ids.dt_analise, self.ids.dt_nota,
                    self.ids.dt_venc,
                    self.ids.num_nota,
                    self.ids.num_cnpj,
                    self.ids.cod_fornec,
                    self.ids.mun_iss,
                    self.ids.regime_trib,
                    self.ids.cod_serv,
                    self.ids.v_bruto,
                    self.ids.aliq_ir,
                    self.ids.irrf,
                    self.ids.aliq_crf,
                    self.ids.crf,
                    self.ids.aliq_inss,
                    self.ids.inss,
                    self.ids.aliq_iss,
                    self.ids.iss,
                    self.ids.v_liq]

        if len(BancoDados.lista) == 0:
            pass
        elif len(BancoDados.lista) == 1:
            BancoDados.lista = BancoDados.lista[0]

            for index, entrada in enumerate(entradas):
                for lista in BancoDados.lista:
                    if index < 10:
                        entrada.text = str(BancoDados.lista[index])
                    else:
                        entrada.text = str(round(float(BancoDados.lista[index]), 2)).replace('.', ',')
            BancoDados.lista.clear()
        else:
            for index, entrada in enumerate(entradas):
                for lista in BancoDados.lista[0]:
                    if index < 10:
                        entrada.text = str(BancoDados.lista[0][index])
                    else:
                        entrada.text = str(round(float(BancoDados.lista[0][index]), 2)).replace('.', ',')
            BancoDados.lista.pop(0)

        print(BancoDados.lista)

    def lembrar_lancamento(self):  # Lembrar informações do último lançamento para notas de mesmo prestador
        if self.ids.lembrar.active:
            lmdb = os.path.join(*self.diretorio, 'Base_notas.accdb;')
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('SELECT TOP 1 data_analise, data, data_vencimento, nf, cnpj, fornecedor, simples_nacional,'
                           'codigo_servico from notas_fiscais order by id desc')
            row = cursor.fetchone()
            self.ids.dt_analise.text = row[0]
            self.ids.dt_nota.text = row[1]
            self.ids.dt_venc.text = row[2]
            self.ids.num_nota.text = str((row[3]) + 1)
            self.ids.num_cnpj.text = row[4]
            self.ids.cod_fornec.text = row[5]
            self.ids.regime_trib.text = row[6]
            self.ids.cod_serv.text = row[7]
            cnx.commit()

        else:
            self.limpar()


class CadastroPrestador(Screen):
    teste = StringProperty(None)

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dialog_cad_err = None
        self.dialog_cad_const = None
        self.dialog_cad = None

    def mascara_cad(self):  # função para formatar CNPJ
        mask = self.ids.cad_cnpj.text
        if mask != '' and '/' not in mask and len(mask) >= 14:
            mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
            self.ids.cad_cnpj.text = mask_cnpj
        else:
            pass

    def pesquisar_prestador(self):  # Pesquisar prestador pelo CNPJ
        try:
            lmdb = os.path.join(*self.manager.get_screen('principal').diretorio, 'Base_notas.accdb;')
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('SELECT * FROM cadastro WHERE CNPJ=?', (self.ids.cad_cnpj.text,))
            row = cursor.fetchone()
            self.ids.cad_cnpj.text = row[0]
            self.ids.cad_nome.text = row[1]
            self.ids.cad_mun.text = row[2]
            self.ids.cad_regime.text = row[3]
            self.ids.aliq_simples.text = str(row[4])
            cnx.commit()
            cnx.close()
        except TypeError:
            self.dialog_cad = MDDialog(text="O CNPJ informado não consta no cadastro!", radius=[20, 7, 20, 7], )
            self.dialog_cad.open()

    def cadastrar_prestador(self):  # Cadastrar novo prestador
        if self.ids.cad_cnpj.text == '':
            pass
        else:
            try:
                lmdb = os.path.join(*self.manager.get_screen('principal').diretorio, 'Base_notas.accdb;')
                cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = cnx.cursor()
                cursor.execute('INSERT INTO cadastro values (?, ?, ?, ?, ?)',
                               (self.ids.cad_cnpj.text, self.ids.cad_nome.text,
                                self.ids.cad_mun.text, self.ids.cad_regime.text, self.ids.aliq_simples.text))
                cnx.commit()
                cnx.close()
                self.dialog_cad = MDDialog(text="Registro incluido com sucesso!", radius=[20, 7, 20, 7], )
                self.dialog_cad.open()
                self.ids.cad_cnpj.text = ''
                self.ids.cad_nome.text = ''
                self.ids.cad_mun.text = ''
                self.ids.cad_regime.text = ''
                self.ids.aliq_simples.text = ''
                self.manager.current = 'principal'

            except pyodbc.DataError:
                self.dialog_cad_err = MDDialog(text="Erro! CNPJ já cadastrado.", radius=[20, 7, 20, 7], )
                self.dialog_cad_err.open()

    def atualizar_cadastro(self):  # Atualizar cadastro após busca pelo CNPJ

        lmdb = os.path.join(*self.manager.get_screen('principal').diretorio, 'Base_notas.accdb;')
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('UPDATE cadastro SET NOME=?, MUNICÍPIO=?, OPTANTE_SIMPLES=?, ALIQUOTA=? WHERE CNPJ=?',
                       (self.ids.cad_nome.text,
                        self.ids.cad_mun.text,
                        self.ids.cad_regime.text,
                        self.ids.aliq_simples.text,
                        self.ids.cad_cnpj.text))
        cnx.commit()
        cnx.close()
        self.dialog_cad_const = MDDialog(text="Cadastro alterado com sucesso!", radius=[20, 7, 20, 7], )
        self.dialog_cad_const.open()


class BancoDados(Screen):
    lista = []

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.total_lancamento = None
        self.data_tables = None

    def gerar_banco(self):  # Gerar banco de dados para visualização
        # conectar banco de dados
        lmdb = os.path.join(*self.manager.get_screen('principal').diretorio, 'Base_notas.accdb;')
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('select * from notas_fiscais order by ID desc')
        resultado = cursor.fetchall()
        cnx.commit()
        cnx.close()
        lin_lancamento = []
        self.total_lancamento = []

        for lin in resultado[:int(self.ids.num_ocor.text)]:  # limitar ultimos lançamentos
            for row in lin:
                if type(row) != str and type(row) != int:
                    lin_lancamento.append(float(row))
                else:
                    lin_lancamento.append(row)
            tupla = tuple(lin_lancamento)

            self.total_lancamento.append(tupla)
            lin_lancamento.clear()

        self.add_datatable()

    def add_datatable(self):  # Adicionar tabela na tela
        self.data_tables = MDDataTable(pos_hint={'center_x': 0.5, 'center_y': 0.5},
                                       size_hint=(1, 0.8),
                                       use_pagination=True, rows_num=10,
                                       background_color_header=get_color_from_hex("#65275d"),
                                       background_color_selected_cell=get_color_from_hex("#eddaeb"),
                                       check=True,
                                       column_data=[("[color=#ffffff]ID[/color]", dp(20)),
                                                    ("[color=#ffffff]Dt_Análise[/color]", dp(20)),
                                                    ("[color=#ffffff]Dt_NF[/color]", dp(20)),
                                                    ("[color=#ffffff]Dt_Venc[/color]", dp(20)),
                                                    ("[color=#ffffff]NF[/color]", dp(20)),
                                                    ("[color=#ffffff]CNPJ[/color]", dp(35)),
                                                    ("[color=#ffffff]Fornecedor[/color]", dp(55)),
                                                    ("[color=#ffffff]Município[/color]", dp(25)),
                                                    ("[color=#ffffff]Regime Trib.[/color]", dp(25)),
                                                    ("[color=#ffffff]Cod. Serv.[/color]", dp(20)),
                                                    ("[color=#ffffff]Val.Bruto[/color]", dp(20)),
                                                    ("[color=#ffffff]Aliq.IR[/color]", dp(15)),
                                                    ("[color=#ffffff]IRRF[/color]", dp(15)),
                                                    ("[color=#ffffff]Aliq.CRF[/color]", dp(15)),
                                                    ("[color=#ffffff]CRF[/color]", dp(15)),
                                                    ("[color=#ffffff]Aliq.INSS[/color]", dp(15)),
                                                    ("[color=#ffffff]INSS[/color]", dp(15)),
                                                    ("[color=#ffffff]Aliq.ISS[/color]", dp(15)),
                                                    ("[color=#ffffff]ISS[/color]", dp(15)),
                                                    ("[color=#ffffff]Val.Líq.[/color]", dp(30)),
                                                    ],
                                       row_data=self.total_lancamento, elevation=1)

        self.add_widget(self.data_tables)

    def pegar_check(self):  # Selecionar notas a serem editadas
        self.lista.clear()
        for item in self.data_tables.get_row_checks():
            self.lista.append(item)


class ExportarDados(Screen):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.dialog_exp = None
        self.dialog = None

    def exp_banco(self):
        # exportar banco completo para consultas e geração de guias de recolhimento
        book = load_workbook(
            os.path.join(*self.manager.get_screen('principal').diretorio, 'Programa Planilha de retenção.xlsx'))
        writer = pd.ExcelWriter(
            os.path.join(*self.manager.get_screen('principal').diretorio, 'Programa Planilha de retenção.xlsx'),
            engine='openpyxl')
        writer.book = book

        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        # Conectar ao banco
        lmdb = os.path.join(*self.manager.get_screen('principal').diretorio, 'Base_notas.accdb;')
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('select * from notas_fiscais')
        resultado = cursor.fetchall()
        lista = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]
        for i in resultado:
            for colunas in range(20):
                lista[colunas].append(i[colunas])
        tabela = pd.DataFrame(lista).transpose()
        tabela.columns = ['ID', 'data_analise', 'data', 'data_vencimento', 'NF', 'CNPJ', 'Fornecedor', 'cidade',
                          'simples_nacional', 'codigo_servico', 'valor_bruto', 'aliq_irrf', 'irrf', 'aliq_crf',
                          'crf',
                          'aliq_inss', 'inss', 'aliq_iss', 'iss', 'valor_liquido']

        cols = ['valor_bruto', 'aliq_irrf', 'irrf', 'aliq_crf', 'crf',
                'aliq_inss', 'inss', 'aliq_iss', 'iss', 'valor_liquido']

        tabela[cols] = tabela[cols].apply(pd.to_numeric, errors='coerce')

        frame = pd.DataFrame(tabela)
        frame.to_excel(writer, sheet_name='Geral', index=False)

        writer.save()
        self.dialog_exp = MDDialog(text="Banco exportado com sucesso!", radius=[20, 7, 20, 7], )
        self.dialog_exp.open()


class Relatorios(Screen):

    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.data_venc = None

    def relatorios(self):
        # Criar planilha para gerar arquivo
        writer = pd.ExcelWriter(os.path.join(*self.manager.get_screen('principal').diretorio, 'Relatórios.xlsx'),
                                engine='xlsxwriter')
        # Conectar ao banco
        lmdb = os.path.join(*self.manager.get_screen('principal').diretorio, 'Base_notas.accdb;')
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()

        # =================== Relatório Imposto de Renda ===============================================#
        if self.ids.check_ir.active:
            cursor.execute('select cnpj, fornecedor, sum(irrf) from notas_fiscais '
                           'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= '
                           'DateValue(?) group by cnpj, fornecedor', self.ids.dt_ini.text, self.ids.dt_fim.text)
            resultado = cursor.fetchall()
            lista = [[], [], []]
            for i in resultado:
                for colunas_ir in range(3):
                    lista[colunas_ir].append(i[colunas_ir])
            tabela = pd.DataFrame(lista).transpose()
            tabela.columns = ['CNPJ', 'Fornecedor', 'IRRF']
            tabela.to_excel(writer, sheet_name='Irrf', index=False)

        # =========================== Relatório de Contribuições =========================================#
        if self.ids.check_crf.active:
            cursor.execute('select * from notas_fiscais where data_vencimento <> 0 (select data_vencimento, '
                           'cnpj, fornecedor, sum(crf) from notas_fiscais '
                           'where DateValue(data_vencimento) >= DateValue(?) and DateValue(data_vencimento) <= '
                           'DateValue(?) '
                           'group by data_vencimento, cnpj, fornecedor order by fornecedor, data_vencimento)',
                           (self.ids.dt_ini.text, self.ids.dt_fim.text))
            resultado = cursor.fetchall()
            lista2 = [[], [], [], []]
            for i in resultado:
                for colunas_crf in range(4):
                    lista2[colunas_crf].append(i[colunas_crf])
            tabela2 = pd.DataFrame(lista2).transpose()
            tabela2.columns = ['Data_Vencimento', 'CNPJ', 'Fornecedor', 'CRF']
            tabela2.to_excel(writer, sheet_name='Crf', index=False)

        # ===========================Relatório ISS ==========================================================#
        # Gerar relatório do ISS por prefeituras em pdf
        if self.ids.check_iss.active:
            # Criar diretório para salvar arquivos pdf
            dir_pdfs = 'ISS_' + self.ids.dt_fim.text[3:].replace('/', '-')
            os.mkdir(os.path.join(*self.manager.get_screen('principal').diretorio, dir_pdfs))
            dados_responsavel = Principal().responsavel.copy()
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()

            cursor.execute('select distinct cidade from notas_fiscais where DateValue(data_analise) >= '
                           'DateValue(?) and DateValue(data_analise) <= DateValue(?)',
                           (self.ids.dt_ini.text, self.ids.dt_fim.text))
            lista = []
            for row in cursor:
                if row[0] != '':
                    lista.append(row[0])

            for i in lista:
                cursor.execute('select NF, fornecedor, iss from notas_fiscais '
                               'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= '
                               'DateValue(?) '
                               'and cidade = ? order by cidade, cnpj', (self.ids.dt_ini.text, self.ids.dt_fim.text, i))

                vencimentos = pd.read_excel(os.path.join(*Principal().diretorio, 'Programa Planilha de retenção.xlsx'),
                                            sheet_name='Relatório ISS', usecols=[9, 10], skiprows=10, dtype=str)

                for index, row in vencimentos.iterrows():
                    if row['MUNICÍPIOS'] == i.upper():
                        dia = vencimentos.loc[index, 'DIA']
                        data = datetime.strptime(self.ids.dt_fim.text, '%d/%m/%Y')
                        data = data + relativedelta(months=1)
                        data = data.strftime('%m/%Y')
                        self.data_venc = dia + '/' + data

                pdf = FPDF(orientation='P', unit='mm', format='A4')
                pdf.add_page()
                pdf.set_font('Arial', 'B', 10)
                pdf.image('logo.png', x=10.0, y=10.0,
                          h=50.0, w=100.0)
                pdf.set_xy(10.0, 70.0)
                pdf.multi_cell(w=125, h=5, txt='ISSQN Município de ' + i)
                pdf.multi_cell(w=125, h=5, txt='A/C: Contabilidade - Contas a Pagar.')
                pdf.set_xy(10.0, pdf.get_y() + 5)
                pdf.multi_cell(w=150, h=5,
                               txt='Planilha contendo valores a recolher referente ao mês ' + self.ids.dt_fim.text[3:])
                pdf.multi_cell(w=125, h=5, txt='Valor a recolher através de BOLETO ANEXO - Contas a Pagar.')
                pdf.multi_cell(w=125, h=5, txt='Vencimento: ' + self.data_venc)
                pdf.set_xy(10.0, pdf.get_y() + 15)
                pdf.multi_cell(w=30, h=5, txt='Nota Fiscal', border=1, align='C')
                pdf.set_xy(40.0, pdf.get_y() - 5)
                pdf.multi_cell(w=80, h=5, txt='Fornecedor', border=1, align='C')
                pdf.set_xy(120.0, pdf.get_y() - 5)
                pdf.multi_cell(w=40, h=5, txt='ISS a recolher', border=1, align='C')
                pdf.set_font('')
                resultado = cursor.fetchall()
                soma = []
                cont = 0
                for lin in resultado:
                    lin[2] = float(lin[2])
                    soma.append(lin[2])
                    lin[2] = str(lin[2]).replace('.', ',')
                    pdf.multi_cell(w=30, h=5, txt=str(lin[0]), border=1, align='C')
                    pdf.set_xy(40.0, pdf.get_y() - 5)
                    pdf.multi_cell(w=80, h=5, txt=str(lin[1][:32]), border=1, align='C')
                    pdf.set_xy(120.0, pdf.get_y() - 5)
                    pdf.multi_cell(w=40, h=5, txt=str(lin[2]), border=1, align='C')
                    cont += 1
                for colunas_ir in range(20 - cont):
                    pdf.multi_cell(w=30, h=5, txt='', border=1)
                    pdf.set_xy(40.0, pdf.get_y() - 5)
                    pdf.multi_cell(w=80, h=5, txt='', border=1)
                    pdf.set_xy(120.0, pdf.get_y() - 5)
                    pdf.multi_cell(w=40, h=5, txt='', border=1)
                pdf.set_xy(40.0, pdf.get_y())
                pdf.multi_cell(w=80, h=5, txt='Valor total a recolher', border=1, align='C')
                pdf.set_xy(120.0, pdf.get_y() - 5)
                pdf.set_font('Arial', 'B', 10)
                pdf.multi_cell(w=40, h=5, txt=str(round(sum(soma), 2)), border=1, align='C')
                pdf.set_xy(10.0, pdf.get_y() + 30)
                pdf.line(10, pdf.get_y(), 60, pdf.get_y())
                pdf.multi_cell(w=100, h=5, txt=dados_responsavel[0])
                pdf.multi_cell(w=100, h=5, txt=dados_responsavel[1])
                pdf.multi_cell(w=40, h=5, txt='GECOT')
                data = self.ids.dt_fim.text[3:].replace('/', '-')
                nome_arquivo = i + ' ' + data + '.pdf'
                pdf.output(os.path.join(*self.manager.get_screen('principal').diretorio, dir_pdfs, nome_arquivo), 'F')

        # ===========================Relatório INSS==========================================================#
        if self.ids.check_inss.active:
            cursor.execute('select data, NF, cnpj, fornecedor, valor_bruto, inss from notas_fiscais '
                           'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) '
                           '<= DateValue(?)',
                           (self.ids.dt_ini.text, self.ids.dt_fim.text))
            resultado = cursor.fetchall()
            lista4 = [[], [], [], [], [], []]
            for i in resultado:
                for colunas_ir in range(6):
                    lista4[colunas_ir].append(i[colunas_ir])
            tabela4 = pd.DataFrame(lista4).transpose()
            tabela4.columns = ['Data Nota Fiscal', 'Nº NF', 'CNPJ', 'Fornecedor', 'Valor Bruto', 'INSS']
            tabela4.to_excel(writer, sheet_name='INSS', index=False)
        else:
            pass

        writer.save()


class WindowManager(ScreenManager):
    pass


class NotasFiscais(MDApp):

    def build(self):
        return Builder.load_file('servicos.kv')


NotasFiscais().run()
