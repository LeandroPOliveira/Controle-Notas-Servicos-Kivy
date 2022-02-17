from kivy import Config
from kivymd.app import MDApp
from kivymd.uix.datatables import MDDataTable
from kivy.lang.builder import Builder
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.metrics import dp
from kivy.core.window import Window


class ContentNavigationDrawer(Screen):
    def mascara(self):
        mask = self.ids.num_cnpj.text
        if mask != '' and '/' not in mask and len(mask) >= 14:
            mask_cnpj = f'{mask[:2]}.{mask[2:5]}.{mask[5:8]}/{mask[8:12]}-{mask[12:14]}'
            self.ids.num_cnpj.text = mask_cnpj
        else:
            pass

class Principal(Screen):
    pass




class Principal2(Screen):
    pass


class WindowManager(ScreenManager):
    pass


class NotasFiscais(MDApp):


    def build(self):
        Window.clearcolor = (1, 1, 1, 1)
        return Builder.load_file('servicos.kv')



NotasFiscais().run()