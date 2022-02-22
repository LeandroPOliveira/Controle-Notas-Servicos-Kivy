from kivy import Config
from kivy.properties import StringProperty, NumericProperty
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
from kivy.core.window import Window
import os
from datetime import date
import pyodbc
from kivy.utils import get_color_from_hex
from kivymd.uix.dialog import MDDialog


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
        if self.ids.regime_trib.text in 'nãoNÃOnaoNAONãoNormalnormal':
            if self.ids.cod_serv.text != '':
                lista = {'irrf': self.ids.aliq_ir, 'crf': self.ids.aliq_crf, 'inss': self.ids.aliq_inss,
                         'iss': self.ids.aliq_iss}
                for imp, aliq in lista.items():
                    cursor.execute(f'select {imp} from tabela_iss where servico = ?', (self.ids.cod_serv.text,))
                    busca = cursor.fetchone()

                    aliq.text = str(round(busca[0], 2)).replace('.', ',')

        cursor.execute(f'select descricao from tabela_iss where servico = ?', (self.ids.cod_serv.text,))
        busca2 = cursor.fetchone()

        self.descr_serv = busca2[0][0:190]

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
        if self.cnpj.get() == '':
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
            # self.lembrar.set(0)
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
            print(i.text)
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

        except:
            self.dialog = MDDialog(text="Erro!", radius=[20, 7, 20, 7], )
            self.dialog.open()



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

    def gerar_banco(self):
        # conectar banco de dados
        lmdb = os.getcwd() + '\Base_notas.accdb;'
        cnx = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'r'DBQ=' + lmdb)
        cursor = cnx.cursor()
        cursor.execute('select * from notas_fiscais order by ID desc')
        self.resultado1 = cursor.fetchall()
        cnx.commit()
        cnx.close()

        self.add_datatable()


    def add_datatable(self):
        self.lista = []
        self.data_tables = MDDataTable(pos_hint={'center_x': 0.5, 'center_y': 0.5},
                                       size_hint=(1, 0.8),
                                       use_pagination=True, rows_num=10,
                                       background_color_header=get_color_from_hex("#03a9e0"),
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
                                                    ("[color=#ffffff]Val.Bruto[/color]", dp(30)),
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
                                       row_data=self.resultado1[:100], elevation=1)

        self.add_widget(self.data_tables)


    def pegar_check(self):
        self.lista.append(self.data_tables.get_row_checks())
        print(self.lista)


class WindowManager(ScreenManager):
    pass


class NotasFiscais(MDApp):
    

    def build(self):
        Window.clearcolor = (1, 1, 1, 1)
        return Builder.load_file('servicos.kv')



NotasFiscais().run()