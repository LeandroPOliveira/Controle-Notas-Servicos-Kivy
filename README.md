# Controle Notas Fiscais de Servico

Objetivos

Mensalmente são emitidas notas fiscais de serviços prestados à empresa por parte de fornecedores diversos, estas notas devem ser analisadas e devem ser apontados os tributos a serem retidos, entre eles IRRF, CRF, INSS e ISS. As notas fiscais são lançadas em planilha excel para controle e posterior emissão das guias e pagamentos dos tributos.

Visando poupar retrabalhos por parte da equipe contábil, este programa tem objetivo de armazenar as informações das notas fiscais em um banco de dados, além de:

Criar uma tabela auxiliar com os códigos de serviços e suas respectivas retenções, a fim de auxiliar o analista em seus trabalhos, esta tabela está disponível no momento em uma API;<br>
Criar um cadastro de fornecedores, utilizando auto-preenchimento no caso de multiplas notas do mesmo emissor;<br>
Para fins de ISS, criar uma tabela com as principais prefeituras e serviços utilizados com retenção, facilitando consultas e eventuais erros de aliquotas na emissão das notas.<br>

Para testar esse script, deve-se comentar as linhas 40, 44, 45 e descomentar as linhas 41, 46 e 47.

---

Goals

Monthly a great number of invoices are issued to the company from multiple suppliers, these invoices must be analyzed, and the taxes should be pointed, in order to pay the government the right amount of taxes. The invoices were filled in a excel file for control and later payments slips.

Aiming to save rework for the accounting team, this program has a goal to store the information of invoices in a simple database, besides:

Create an auxiliary table with service codes and their respective tax retention, to support the analyst on his work, this table is now on an API;<br>
Create a supplier registration, using and autofill in case of multiple invoices of same supplier;<br>
For a specific tax of City Halls, create a table with the main services most used, facilitating queries and eventual mistakes of aliquot on issued invoices.<br>

To test this script, you must comment the lines 40, 44, 45 e uncomment the lines 41, 46 e 47.

 
---

Tela Principal / Main Screen:<br>

<img src="https://github.com/LeandroPOliveira/Controle-Notas-Servicos-Kivy/blob/main/cadastro-nota.gif" width="800" height="400"><br><br>

Tela Consultas / Query Screen:<br>

<img src="https://github.com/LeandroPOliveira/Controle-Notas-Servicos-Kivy/blob/main/consulta-nota.gif" width="800" height="400">
