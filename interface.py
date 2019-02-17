from tkinter import ttk
from tkinter import *
from tinydb import TinyDB, Query
import os
import main #Importa modulo que opera Excel

class Cliente:
	""" Método inicial que maqueta la interfaz con el modulo Tkinter e instancia el resto métodos """
	def __init__(self, window):
		self.wind = window
		self.wind.title('Clientes MantrepHH')

		#Creating a Frame Container
		frame = LabelFrame(self.wind, text ='Registrar Nuevo Cliente')
		frame.grid(row = 0, column = 0, columnspan = 3)

		#Client Input
		Label(frame, text='Cliente: ').grid(row = 1, column = 0, sticky = 'E')
		self.client = Entry(frame)
		self.client.focus()
		self.client.config(width = 50)
		self.client.grid(row = 1, column = 1)

		#Representative Input
		Label(frame, text='Representante: ').grid(row = 2, column = 0, sticky = 'E')
		self.representative = Entry(frame)
		self.representative.config(width = 50)
		self.representative.grid(row = 2, column = 1)

		#Address Input
		Label(frame, text='Dirección: ').grid(row = 3, column = 0, sticky = 'E')
		self.address = Entry(frame)
		self.address.config(width = 50)
		self.address.grid(row = 3, column = 1)

		#City Input
		Label(frame, text='Ciudad: ').grid(row = 4, column = 0, sticky = 'E')
		defaultCity = StringVar(frame, value = 'Manizales')
		self.city = Entry(frame, textvariable = defaultCity)
		self.city.config(width = 50)
		self.city.grid(row = 4, column = 1)

		#Phone Input
		Label(frame, text='Teléfono: ').grid(row = 5, column = 0, sticky = 'E')
		self.phone = Entry(frame)
		self.phone.config(width = 50)
		self.phone.grid(row = 5, column = 1)

		#Email Input
		Label(frame, text='Email: ').grid(row = 6, column = 0, sticky = 'E')
		self.email = Entry(frame)
		self.email.config(width = 50)
		self.email.grid(row = 6, column = 1)

		#Button para agregar cliente a BD
		ttk.Button(frame, text='Agregar Cliente', command = self.addClient).grid(row = 7, columnspan = 2, sticky = W + E)

		#Mensajes de confirmación
		self.message = Label(text = '', fg = 'green')
		self.message.grid(row = 7, column = 0, columnspan = 3, sticky = W + E)

		#Table
		self.tree = ttk.Treeview(height = 10, columns = ('#0','#1','#2'))
		self.tree.grid(row = 9, column = 0, columnspan = 3)
		self.tree.column('#0', width = 200)
		self.tree.column('#1', width = 200)
		self.tree.column('#2', width = 100)
		self.tree.column('#3', width = 150)
		self.tree.heading('#0', text = 'Nombre', anchor = CENTER)
		self.tree.heading('#1', text = 'Representante', anchor = CENTER)
		self.tree.heading('#2', text = 'Telefono', anchor = CENTER)
		self.tree.heading('#3', text = 'Email', anchor = CENTER)

		#Button Editar y Eliminar
		ttk.Button(text = 'EDITAR', command = self.editClient).grid(row = 8, column = 0, sticky = W + E)
		ttk.Button(text = 'ELIMINAR', command = self.deleteClient).grid(row = 8, column = 1, sticky = W + E)
		ttk.Button(text = 'NUEVA COTIZACION', command = self.createCot).grid(row = 8, column = 2, sticky = W + E, pady = 10)	

		#Llenando las filas de la tabla
		self.getClients()

	"""
		Método que se ejecuta para actualizar la tabla TreeView:
		1. Elimina cada fila en la tabla
		2. Consulta la BD
		3. Llena la tabla con los datos obtenidos
	"""
	def getClients(self, parameters = ()):
		db = TinyDB('db/db.json') #Guarda en variable la BD
		clients = db.all() #Consulta y guarda todos los datos de la BD

		#Limpia los datos existentes en la tabla
		records = self.tree.get_children()
		for element in records:
			self.tree.delete(element)

		#Obtiene los datos de la BD y los guarda en la tabla
		for client in clients:
			self.tree.insert('', 0, text = client["Client"], values = (client["Name"], client["Phone"], client["Email"]))

	""" Método que retorna booleano al verificar cada entrada por teclado del usuario al registrar nuevo cliente """
	def validate(self):
		inputCliente = len(self.client.get()) != 0
		inputNombre = len(self.representative.get()) != 0
		inputDir = len(self.address.get()) != 0
		inputCiudad = len(self.city.get()) != 0
		inputTel = len(self.phone.get()) > 6
		inputEmail = len(self.email.get()) != 0 and self.email.get().find('@') > 0 and self.email.get().find('.') != -1

		return inputCliente and inputNombre and inputDir and inputCiudad and inputTel and inputEmail

	""" Método usado para limpiar mensajes existentes """
	def clearMsg(self):
		self.message['text'] = ''

	######## SECCION CRUD ########
	
	""" Método que inserta los datos, ya verificados, en la BD y actualiza la tabla """
	def addClient(self):
		#Si la validacion retorna True...
		if self.validate():
			#Conexión a BD
			db = TinyDB('db/db.json')
			Cliente = Query()

			client_name = self.client.get()
			client_name = client_name.upper() #Convierte el string en MAYUSCULAS

			#Inserta los datos usando db.insert en el orden
			db.insert({'Client': client_name, 'Count': 0, 'Name' : self.representative.get(), 'Address' : self.address.get(), 'City' : self.city.get(), 'Phone' : self.phone.get(), 'Email' : self.email.get()})

			#Genera un nuevo mensaje en ventana principal y lo oculta luego de 3500ms
			self.message['fg'] = 'green'
			self.message['text'] = "Cliente {} añadido exitosamente".format(client_name)
			self.wind.after(3500, self.clearMsg) #Oculta el mensaje con el método clearMsg

			#Crea carpeta que almacenara las cotizaciones
			os.system('mkdir "..\\%s"' % client_name)

			#Limpia el formulario 
			self.client.delete(0, END)
			self.representative.delete(0, END)
			self.address.delete(0, END)
			self.phone.delete(0, END)
			self.email.delete(0, END)
		#Si la validación retorna False lanza mensaje y lo oculta luego de 2500ms
		else:
			self.message['fg'] = 'red'
			self.message['text'] = "Todos los datos son obligatorios"
			self.wind.after(2500, self.clearMsg)

		#Actualiza la tabla en ventana principal
		self.getClients()

	""" Método que elimina al cliente y su conjunto de datos """
	def deleteClient(self):
		#Conexión a BD
		db = TinyDB('db/db.json')
		Cliente = Query()
		
		#Verifica que haya un elemento seleccionado de la tabla para proceder
		try:
			self.tree.item(self.tree.selection())['text'][0]
		except IndexError as e:
			self.message['fg'] = 'red'
			self.message['text'] = 'Para ELIMINAR debes seleccionar un cliente'
			self.wind.after(2500, self.clearMsg)
			return

		name = self.tree.item(self.tree.selection())['text'] #Obtiene el nombre seleccionado en tabla y lo guarda en 'name'
		db.remove(Cliente.Client == name) #Elimina los datos pertenecientes al elemento seleccionado

		#Imprime mensaje y lo elimina luego de 3500ms
		self.message['fg'] = 'red'
		self.message['text'] = 'Cliente {} eliminado exitosamente'.format(name)
		self.wind.after(3500, self.clearMsg)

		#Actualiza la tabla
		self.getClients()

	""" Método que ejecuta nueva ventana con los datos del elemento seleccionado para editarlos """
	def editClient(self):
		#Conexión a BD
		db = TinyDB('db/db.json')
		Cliente = Query()
		
		#Verifica que haya un elemento seleccionado de la tabla para proceder
		try:
			self.tree.item(self.tree.selection())['text'][0]
		except IndexError as e:
			self.message['fg'] = 'red'
			self.message['text'] = 'Para EDITAR debes seleccionar un cliente'
			self.wind.after(2500, self.clearMsg)
			return

		name = self.tree.item(self.tree.selection())['text'] #Obtiene el nombre seleccionado en tabla y lo guarda en 'name'
		queryClient = db.search(Cliente.Client == name) #Obtiene todos los datos pertenecientes y los guarda

		#Guarda cada dato del cliente seleccionado en variables independientes
		RepCli = queryClient[0]['Name']
		AddCli = queryClient[0]['Address']
		CityCli = queryClient[0]['City']
		PhoneCli = queryClient[0]['Phone']
		EmailCli = queryClient[0]['Email']

		#Crea nueva ventana con el metodo TopLevel()
		self.editView = Toplevel()
		self.editView.title('Editar Cliente')

		#Formulario con los datos existentes
		Label(self.editView, text = 'Cliente: ').grid(row = 0, column = 1, sticky = 'E')
		c = Entry(self.editView, textvariable = StringVar(self.editView, value = name), state = 'readonly')
		c.grid(row = 0, column = 2)
		c.config(width = 50)

		Label(self.editView, text = 'Representante: ').grid(row = 1, column = 1, sticky = 'E')
		varRep = Entry(self.editView, textvariable = StringVar(self.editView, value = RepCli))
		varRep.grid(row = 1, column = 2)
		varRep.config(width = 50)

		Label(self.editView, text = 'Dirección: ').grid(row = 2, column = 1, sticky = 'E')
		varAdd = Entry(self.editView, textvariable = StringVar(self.editView, value = AddCli))
		varAdd.grid(row = 2, column = 2)
		varAdd.config(width = 50)

		Label(self.editView, text = 'Ciudad: ').grid(row = 3, column = 1, sticky = 'E')
		varCity = Entry(self.editView, textvariable = StringVar(self.editView, value = CityCli))
		varCity.grid(row = 3, column = 2)
		varCity.config(width = 50)

		Label(self.editView, text = 'Teléfono: ').grid(row = 4, column = 1, sticky = 'E')
		varPhone = Entry(self.editView, textvariable = StringVar(self.editView, value = PhoneCli))
		varPhone.grid(row = 4, column = 2)
		varPhone.config(width = 50)

		Label(self.editView, text = 'Email: ').grid(row = 5, column = 1, sticky = 'E')
		varEmail = Entry(self.editView, textvariable = StringVar(self.editView, value = EmailCli))
		varEmail.grid(row = 5, column = 2)
		varEmail.config(width = 50)

		#Actualiza los datos nuevos cuando usa el método updateClient
		Button(self.editView, text = 'Actualizar', command = lambda: self.updateClient(name, varRep.get(), varAdd.get(), varCity.get(), varPhone.get(), varEmail.get())).grid(row = 6, column = 1)

	""" Recibe los datos de la ventada para editar cliente y los actualiza en la BD """
	def updateClient(self,base, rep, add, city, phone, email):
		#Conexión a BD
		db = TinyDB('db/db.json')
		Cliente = Query()

		#Actualiza los datos pasados por la ventana EditView
		db.update({"Name" : rep, "Address" : add, "City" : city, "Phone": phone, "Email" : email}, Cliente.Client == base)
		
		#Destruye la venta EditView
		self.editView.destroy()

		#Mensaje de confirmacion en ventana principal
		self.message['fg'] = 'green'
		self.message['text'] = 'Cliente {} editado exitosamente'.format(base)
		self.wind.after(3500, self.clearMsg)

		#Actualiza TreeView
		self.getClients()

	""" Método que verifica el elemento a usar para crear nueva cotizacion """
	def createCot(self):
		#Conexión a BD
		db = TinyDB('db/db.json')
		Cliente = Query()
		
		#Verifica que haya un elemento seleccionado de la tabla para proceder
		try:
			self.tree.item(self.tree.selection())['text'][0]
		except IndexError as e:
			self.message['fg'] = 'red'
			self.message['text'] = 'Para crear una cotizacion debes seleccionar un cliente'
			self.wind.after(2500, self.clearMsg)
			return

		name = self.tree.item(self.tree.selection())['text'] #Guarda el nombre del cliente seleccionado
		main.CrearCotizacion(name) #Usa el modulo main para crear la cotización

		#Imprime mensaje de confirmación y lo oculta luego de 3500ms
		self.message['fg'] = 'green'
		self.message['text'] = 'Cotizacion creada existosamente para {}'.format(name)
		self.wind.after(3500, self.clearMsg)

if __name__ == '__main__':
	window = Tk()
	app = Cliente(window)
	window.mainloop()
