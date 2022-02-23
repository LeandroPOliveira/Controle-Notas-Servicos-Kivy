from kivy.properties import StringProperty
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
from kivy.core.window import Window
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
    # cod_id = StringProperty(0)

    def mascara(self):
        mask = self.ids.num_cnpj.text
        if mask != '' and '/' not in mask and len(mask) >= 14:
            mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
            self.ids.num_cnpj.text = mask_cnpj
        else:
            pass

    def busca_cadastro(self):
        if self.ids.num_cnpj.text != '':
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = self.cnx.cursor()
            cursor.execute('select nome from cadastro where cnpj = ?', (self.ids.num_cnpj.text,))
            busca_nome = cursor.fetchone()

            self.ids.cod_fornec.text = busca_nome[0]

            cursor.execute('select optante_simples from cadastro where cnpj = ?', (self.ids.num_cnpj.text,))
            busca_simples = cursor.fetchone()
            self.ids.regime_trib.text = busca_simples[0]

    def busca_servico(self):
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        self.cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = self.cnx.cursor()
        if self.ids.cod_serv.text != '':
            lista = {'irrf': self.ids.aliq_ir, 'crf': self.ids.aliq_crf, 'inss': self.ids.aliq_inss,
                     'iss': self.ids.aliq_iss}
            for imp, aliq in lista.items():
                if self.ids.regime_trib.text in 'nãoNÃOnaoNAONãoNormalnormal':
                    cursor.execute(f'select {imp} from tabela_iss where servico = ?', (self.ids.cod_serv.text,))
                    busca = cursor.fetchone()
                    aliq.text = str(round(busca[0], 2)).replace('.', ',')
                else:
                    aliq.text = '0,00'

        if self.ids.regime_trib.text not in 'Simplessimples':
            cursor.execute('select aliq_iss from municipios where municipio = ? and cod_iss = ?',
                           (self.ids.mun_iss.text, self.ids.cod_serv.text, ))
            busca_aliq = cursor.fetchone()
            self.ids.aliq_iss.text = str(round(busca_aliq[0], 2)).replace('.', ',')

        try:
            cursor.execute(f'select descricao from tabela_iss where servico = ?', (self.ids.cod_serv.text,))
            busca2 = cursor.fetchone()
            self.descr_serv = busca2[0][0:190]
        except:
            pass

    def calcula_imposto(self, instance, aliquota):
        if aliquota.text != '':
            tupla = (aliquota.text.replace(',', '.'), self.ids.v_bruto.text.replace(',', '.'))
            instance.text = str(round(float(tupla[1]) * (float(tupla[0]) / 100), 2)).replace('.', ',')
        else:
            instance.text = '0'
            aliquota.text = '0'


    def valor_liq(self):
        self.ids.v_liq.text = str(round(float(self.ids.v_bruto.text.replace(',', '.')) -
         (sum([float(self.ids.irrf.text.replace(',', '.')),
         float(self.ids.crf.text.replace(',', '.')),
         float(self.ids.inss.text.replace(',', '.')),
         float(self.ids.iss.text.replace(',', '.'))])), 2)).replace('.', ',')

    def data_dia(self):
        if self.ids.dt_nota.text == '':
            self.ids.dt_analise.text = date.today().strftime('%d/%m/%Y')
        else:
            pass




    def adicionar(self):
        if self.ids.num_cnpj.text == '':
            self.dialog = MDDialog(
                text="Insira todas as informações!",
                radius=[20, 7, 20, 7], )

            self.dialog.open()
        else:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
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
                                 self.ids.v_liq.text))
            self.ids.lembrar.active = False
            self.limpar()
            cnx.commit()
            cnx.close()

            self.dialog = MDDialog(text="Registro incluido com sucesso!", radius=[20, 7, 20, 7], )
            self.dialog.open()



    def limpar(self):
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

    def apagar(self):
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('DELETE FROM notas_fiscais WHERE ID=?', (self.ids.cod_id.text,))
        cnx.commit()
        cnx.close()
        self.dialog = MDDialog(text="Registro apagado com sucesso!", radius=[20, 7, 20, 7], )
        self.dialog.open()
        self.limpar()

    def buscar(self):
        try:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
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
            self.ids.v_bruto.text = str(round(row[10],2)).replace('.',',')
            self.ids.aliq_ir.text = str(round(row[11],2)).replace('.',',')
            self.ids.irrf.text = str(round(row[12],2)).replace('.',',')
            self.ids.aliq_crf.text = str(round(row[13],2)).replace('.',',')
            self.ids.crf.text = str(round(row[14],2)).replace('.',',')
            self.ids.aliq_inss.text = str(round(row[15],2)).replace('.',',')
            self.ids.inss.text = str(round(row[16],2)).replace('.',',')
            self.ids.aliq_iss.text = str(round(row[17],2)).replace('.',',')
            self.ids.iss.text = str(round(row[18],2)).replace('.',',')
            self.ids.v_liq.text = str(round(row[19],2)).replace('.',',')
            cnx.commit()
            cnx.close()
        except:
            self.dialog = MDDialog(text="Registro não encontrado!", radius=[20, 7, 20, 7], )
            self.dialog.open()
            self.limpar()


    def atualizar(self):
        try:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('update notas_fiscais set DATA_ANALISE=?, DATA=?, DATA_VENCIMENTO=?, NF=?, CNPJ=?, FORNECEDOR=?, '
                           'CIDADE=?, SIMPLES_NACIONAL=?, CODIGO_SERVICO=?, VALOR_BRUTO=?, ALIQ_IRRF=?, IRRF=?, ALIQ_CRF=?, '
                           'crf=?, ALIQ_INSS=?, INSS=?, ALIQ_ISS=?, ISS=?, VALOR_LIQUIDO=? where ID=?',(self.ids.dt_analise.text,
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
            self.dialog = MDDialog(text="Registro alterado com sucesso!", radius=[20, 7, 20, 7], )
            self.dialog.open()
            self.limpar()
            self.inserir_notas()

        except:
            self.dialog = MDDialog(text="Erro!", radius=[20, 7, 20, 7], )
            self.dialog.open()



    def inserir_notas(self):

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
        #
        # print(BancoDados.lista)
        # print(len(BancoDados.lista))

        if len(BancoDados.lista) == 0:
            pass
        elif len(BancoDados.lista) == 1:
            BancoDados.lista = BancoDados.lista[0]

            for index, entrada in enumerate(entradas):
                for lista in BancoDados.lista:
                    if index < 10:
                        entrada.text = str(BancoDados.lista[index])
                    else:
                        entrada.text = str(round(float(BancoDados.lista[index]),2)).replace('.', ',')
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

    def lembrar_lancamento(self):
        if self.ids.lembrar.active == True:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
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

    def mascara_cad(self):  # função para formatar CNPJ
        mask = self.ids.cad_cnpj.text
        if mask != '' and '/' not in mask and len(mask) >= 14:
            mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
            self.ids.cad_cnpj.text = mask_cnpj
        else:
            pass

    def pesquisar_fornecedor(self):
        try:
            lmdb = os.getcwd() + '\Base_notas.accdb;'
            cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
            cursor = cnx.cursor()
            cursor.execute('SELECT * FROM cadastro WHERE CNPJ=?', (self.ids.cad_cnpj.text,))
            row = cursor.fetchone()
            self.ids.cad_cnpj.text = row[0]
            self.ids.cad_nome.text = row[1]
            self.ids.cad_mun.text = row[2]
            self.ids.cad_regime.text = row[3]
            cnx.commit()
            cnx.close()
        except:
            self.dialog = MDDialog(
                text="O CNPJ informado não consta no cadastro!",
                radius=[20, 7, 20, 7],)

            self.dialog.open()


    def cadastrar_prestador(self):
        if self.ids.cad_cnpj.text == '':
            pass
            # tkinter.messagebox.showerror('Notas fiscais de Serviço', 'Coloque todas as informações')
        else:
            try:
                lmdb = os.getcwd() + '\Base_notas.accdb;'
                cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
                cursor = cnx.cursor()
                cursor.execute('INSERT INTO cadastro values (?, ?, ?, ?)', (self.ids.cad_cnpj.text, self.ids.cad_nome.text,
                                                                            self.ids.cad_mun.text, self.ids.cad_regime.text))
                cnx.commit()
                # tkinter.messagebox.showinfo('Notas Fiscais de Serviço', 'Registro incluído com sucesso!')
                cnx.close()

            except:
                pass
                # tkinter.messagebox.showerror('Notas Fiscais de Serviço', 'Erro! CNPJ já cadastrado!')


    def atualizar_cadastro(self):

        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('UPDATE cadastro SET NOME=?, MUNICÍPIO=?, OPTANTE_SIMPLES=? WHERE CNPJ=?',
                       (self.ids.cad_nome.text,
                        self.ids.cad_mun.text,
                        self.ids.cad_regime.text,
                        self.ids.cad_cnpj.text))
        cnx.commit()
        cnx.close()




class BancoDados(Screen):
    lista = []

    def gerar_banco(self):
        # conectar banco de dados
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('select * from notas_fiscais order by ID desc')
        resultado = cursor.fetchall()
        cnx.commit()
        cnx.close()
        lin_lancamento = []
        self.total_lancamento = []

        for lin in resultado[:100]:
            for row in lin:
                if type(row) != str and type(row) != int:
                    lin_lancamento.append(float(row))
                else:
                    lin_lancamento.append(row)
            tupla = tuple(lin_lancamento)

            self.total_lancamento.append(tupla)
            lin_lancamento.clear()




        self.add_datatable()


    def add_datatable(self):

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


    def pegar_check(self):
        self.lista.clear()
        for item in self.data_tables.get_row_checks():
            self.lista.append(item)

class ExportarDados(Screen):


    def exp_banco(self):
        # exportar banco completo para consultas e geração de guias de recolhimento
        book = load_workbook('Programa Planilha de retenção.xlsx')
        writer = pd.ExcelWriter('Programa Planilha de retenção.xlsx', engine='openpyxl')
        writer.book = book

        writer.sheets = dict((ws.title, ws) for ws in book.worksheets)

        # Conectar ao banco
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('select * from notas_fiscais')
        resultado = cursor.fetchall()
        lista = [[], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], [], []]
        for i in resultado:
            for l in range(20):
                lista[l].append(i[l])
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
        self.dialog = MDDialog(text="Banco exportado com sucesso!", radius=[20, 7, 20, 7], )
        self.dialog.open()

class Relatorios(Screen):


    def relatorios(self):

        # Criar planilha para gerar arquivo
        writer = pd.ExcelWriter(self.ids.diretorio.text + '\Relatórios.xlsx', engine='xlsxwriter')
        # Conectar ao banco
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()

        # =================== Relatório Imposto de Renda ===============================================#
        if self.ids.check_ir.active == True:
            cursor.execute('select cnpj, fornecedor, sum(irrf) from notas_fiscais '
                           'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= '
                           'DateValue(?) group by cnpj, fornecedor', self.ids.dt_ini.text, self.ids.dt_fim.text)
            resultado = cursor.fetchall()
            lista = [[], [], []]
            for i in resultado:
                for l in range(3):
                    lista[l].append(i[l])
            tabela = pd.DataFrame(lista).transpose()
            tabela.columns = ['CNPJ', 'Fornecedor', 'IRRF']
            tabela.to_excel(writer, sheet_name='Irrf', index=False)

        # =========================== Relatório de Contribuições =========================================#
        if self.ids.check_crf.active == True:
            cursor.execute('select * from notas_fiscais where data_vencimento <> 0 (select data_vencimento, '
                           'cnpj, fornecedor, sum(crf) from notas_fiscais '
                           'where DateValue(data_vencimento) >= DateValue(?) and DateValue(data_vencimento) <= DateValue(?) '
                           'group by data_vencimento, cnpj, fornecedor order by fornecedor, data_vencimento)',
                           (self.ids.dt_ini.text, self.ids.dt_fim.text))
            resultado = cursor.fetchall()
            lista2 = [[], [], [], []]
            for i in resultado:
                for l in range(4):
                    lista2[l].append(i[l])
            tabela2 = pd.DataFrame(lista2).transpose()
            tabela2.columns = ['Data_Vencimento', 'CNPJ', 'Fornecedor', 'CRF']
            tabela2.to_excel(writer, sheet_name='Crf', index=False)

        # ===========================Relatório ISS ==========================================================#
        if self.ids.check_iss.active == True:
            lmdb = os.getcwd() + '\\base_notas.accdb;'
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
                               'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) <= DateValue(?)'
                               'and cidade = ? order by cidade, cnpj', (self.ids.dt_ini.text, self.ids.dt_fim.text, i))

                vencimentos = pd.read_excel('G:\GECOT\FISCAL - Retenções\\Programa Planilha de retenção.xlsx',
                                            sheet_name='Relatório ISS', usecols=[9, 10], skiprows=10, dtype=str)

                for index, row in vencimentos.iterrows():
                    if row['MUNICÍPIOS'] == i.upper():
                        dia = vencimentos.loc[index, 'DIA']
                        data = datetime.strptime(self.ids.dt_fim.text, '%d/%m/%Y')
                        data = data + relativedelta(months=1)
                        data = data.strftime('%m/%Y')
                        data_venc = dia + '/' + data

                pdf = FPDF(orientation='P', unit='mm', format='A4')
                pdf.add_page()
                pdf_w = 210
                pdf_h = 297
                pdf.set_font('Arial', 'B', 10)
                pdf.image('G:\GECOT\FISCAL - Retenções\logo.png', x=10.0, y=10.0,
                          h=50.0, w=100.0)
                pdf.set_xy(10.0, 70.0)
                pdf.multi_cell(w=125, h=5, txt='ISSQN Município de ' + i)
                pdf.multi_cell(w=125, h=5, txt='A/C: Contabilidade - Contas a Pagar.')
                pdf.set_xy(10.0, pdf.get_y() + 5)
                pdf.multi_cell(w=150, h=5,
                               txt='Planilha contendo valores a recolher referente ao mês ' + self.ids.dt_fim.text[3:])
                pdf.multi_cell(w=125, h=5, txt='Valor a recolher através de BOLETO ANEXO - Contas a Pagar.')
                pdf.multi_cell(w=125, h=5, txt='Vencimento: ' + data_venc)
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
                for l in range(20 - cont):
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
                pdf.multi_cell(w=100, h=5, txt='Pedro Henrique Carrilho')
                pdf.multi_cell(w=100, h=5, txt='Contador Junior')
                pdf.multi_cell(w=40, h=5, txt='GECOT')
                data = self.ids.dt_fim.text[3:].replace('/', '-')
                pdf.output(i + ' ' + data + '.pdf', 'F')

        # ===========================Relatório INSS==========================================================#
        if self.ids.check_inss.active == True:
            cursor.execute('select data, NF, cnpj, fornecedor, valor_bruto, inss from notas_fiscais '
                           'where DateValue(data_analise) >= DateValue(?) and DateValue(data_analise) '
                           '<= DateValue(?)',
                           (self.ids.dt_ini.text, self.ids.dt_fim.text))
            resultado = cursor.fetchall()
            lista4 = [[], [], [], [], [], []]
            for i in resultado:
                for l in range(6):
                    lista4[l].append(i[l])
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
        Window.clearcolor = (1, 1, 1, 1)
        return Builder.load_file('servicos.kv')



NotasFiscais().run()