from kivy import Config
from kivy.properties import StringProperty
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

class ContentNavigationDrawer(Screen):
    pass

class Principal(Screen):
    descr_serv = StringProperty('')


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


class Principal2(Screen):
    pass

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


class WindowManager(ScreenManager):
    pass


class NotasFiscais(MDApp):


    def build(self):
        Window.clearcolor = (1, 1, 1, 1)
        return Builder.load_file('servicos.kv')



NotasFiscais().run()