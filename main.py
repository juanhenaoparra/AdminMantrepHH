from openpyxl import load_workbook
from tinydb import TinyDB, Query
import os

""" Método que suma el contador de archivos e instancia el siguiente metodo para crear el archivo de excel """
def CrearCotizacion(client_name):
	#Conexión a BD
	db = TinyDB('db/db.json')
	Cliente = Query()

	#Guarda los elementos que pertenecen al nombre pasado por parametro
	queryClient = db.search(Cliente.Client == client_name)
	Count = int(queryClient[0]['Count'])
	newCount = Count+1 #Le suma 1 al contador

	db.update({'Count':newCount}, Cliente.Client == client_name) #Actualiza el contador en la BD
	os.system('copy HojadeCotizacion.xlsx "..\\%s"' % client_name) #Copia el archivo en la carpeta respectiva del cliente

	#Condicional compuesto que renombra el archivo segun su código
	if newCount <= 9:
		os.rename('..//'+client_name+'//HojadeCotizacion.xlsx', '..//'+client_name+'//C-0000'+str(newCount)+'.xlsx')
	elif newCount >= 10 and newCount <= 99:
		os.rename('..//'+client_name+'//HojadeCotizacion.xlsx', '..//'+client_name+'//C-000'+str(newCount)+'.xlsx')
	elif newCount >= 100 and newCount <= 999:
		os.rename('..//'+client_name+'//HojadeCotizacion.xlsx', '..//'+client_name+'//C-00'+str(newCount)+'.xlsx')
	elif newCount >= 1000 and newCount <= 9999:
		os.rename('..//'+client_name+'//HojadeCotizacion.xlsx', '..//'+client_name+'//C-0'+str(newCount)+'.xlsx')
	elif newCount >= 10000 and newCount <= 99999:
		os.rename('..//'+client_name+'//HojadeCotizacion.xlsx', '..//'+client_name+'//C-'+str(newCount)+'.xlsx')

	#Llama la funcion para manipular el archivo creado
	ManipularCotizacion(client_name)

""" Método que modifica y abre la cotizacion en Excel """
def ManipularCotizacion(client_name):
	#Conexión a BD
	db = TinyDB('db/db.json')
	Cliente = Query()

	#Guarda los elementos que pertenecen al nombre pasado por parametro
	queryClient = db.search(Cliente.Client == client_name)
	RepCli = queryClient[0]['Name']
	AddCli = queryClient[0]['Address']
	CityCli = queryClient[0]['City']
	PhoneCli = queryClient[0]['Phone']
	EmailCli = queryClient[0]['Email']
	Count = int(queryClient[0]['Count'])

	nCot = '' #String vacío que almacenará el codigo del archivo

	#Condicional que guarda en 'nCot' el código
	if Count <= 9:
		nCot = 'C-0000'+str(Count)
	elif Count >= 10 and Count <= 99:
		nCot = 'C-000'+str(Count)
	elif Count >= 100 and Count <= 999:
		nCot = 'C-00'+str(Count)
	elif Count >= 1000 and Count <= 9999:
		nCot = 'C-0'+str(Count)
	elif Count >= 10000 and Count <= 99999:
		nCot = 'C-'+str(Count)

	#Ruta donde se ubica el archivo del cliente
	FILE_PATH = '..//'+client_name+'//'+nCot+'.xlsx'
	SHEET = 'Cotizacion'

	#Se usa el modulo openpyxl
	workbook = load_workbook(FILE_PATH)
	sheet = workbook[SHEET]
	
	#Se modifican las celdas con los datos del cliente en la BD
	sheet['B4'] = client_name
	sheet['B6'] = nCot
	sheet['B10'] = AddCli
	sheet['B12'] = CityCli
	sheet['B14'] = PhoneCli
	sheet['B16'] = EmailCli
	sheet['B18'] = RepCli
	
	#Se guardan los cambios hechos
	workbook.save(FILE_PATH)

	#Se abre el archivo
	comandoInicio = 'start ..//"%s"//%s.xlsx' % (client_name, nCot)
	os.system(comandoInicio)
