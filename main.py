from openpyxl import load_workbook
from tinydb import TinyDB, Query
import os

def CrearCotizacion(client_name):
	db = TinyDB('db/db.json')
	Cliente = Query()
	queryClient = db.search(Cliente.NombreCliente == client_name)
	Count = int(queryClient[0]['Count'])
	newCount = Count+1
	db.update({'Count':newCount}, Cliente.NombreCliente == client_name)
	os.system('copy HojadeCotizacion.xlsx "..\\%s"' % client_name)

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

def ValidarNombreCliente(client_name):
	db = TinyDB('db/db.json')
	Cliente = Query()
	clienteExistente = False

	for item in db:
		NombreCliente = item['NombreCliente']
		if NombreCliente == client_name:
			clienteExistente = True

	if not clienteExistente:
		print("Creando nuevo cliente...")
		db.insert({'NombreCliente': client_name, 'Count': 1})
		os.system('mkdir "..\\%s"' % client_name)
		os.system('copy HojadeCotizacion.xlsx "..\\%s"' % client_name)
		os.rename('..//'+client_name+'//HojadeCotizacion.xlsx', '..//'+client_name+'//C-00001.xlsx')

	else:
		queryClient = db.search(Cliente.NombreCliente == client_name)
		print('Creando nueva cotizacion para '+queryClient[0]['NombreCliente'])
		CrearCotizacion(client_name)

def AbrirCotizacion(client_name):
	db = TinyDB('db/db.json')
	Cliente = Query()
	queryClient = db.search(Cliente.NombreCliente == client_name)
	Count = int(queryClient[0]['Count'])
	nCot = ''

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

	FILE_PATH = '..//'+client_name+'//'+nCot+'.xlsx'
	SHEET = 'Cotizacion'

	workbook = load_workbook(FILE_PATH)
	sheet = workbook[SHEET]
	#b6 = sheet['B6'].value	#Numero de Cotización [Formato: C-00001]
	sheet['B4'] = client_name
	sheet['B6'] = nCot
	workbook.save(FILE_PATH)

	comandoInicio = 'start ..//"%s"//%s.xlsx' % (client_name, nCot)
	os.system(comandoInicio)


print("----MANTENIMIENTO Y REPARACIONES HH----\n------------NIT 9009704435-------------\n")

while True:
	client_name = input("\nIngrese el nombre del cliente: ")
	if client_name == '0':
		break
	client_name = client_name.upper()

	ValidarNombreCliente(client_name)
	AbrirCotizacion(client_name)
