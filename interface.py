from tkinter import ttk
from tkinter import *
from tinydb import TinyDB, Query
import main
import os

class Cliente:
	def __init__(self, window):
		self.wind = window
		self.wind.title('Clientes MantrepHH')

		#Creating a Frame Container
		frame = LabelFrame(self.wind, text='Registrar Nuevo Cliente')
		frame.grid(row = 0, column = 0, columnspan= 3, pady = 20)

		#Client Input
		Label(frame, text='Cliente: ').grid(row = 1, column = 0)
		self.client = Entry(frame)
		self.client.focus()
		self.client.grid(row = 1, column = 1)

		#Representative Input
		Label(frame, text='Nombre Representante: ').grid(row = 2, column = 0)
		self.representative = Entry(frame)
		self.representative.grid(row = 2, column = 1)

		#Address Input
		Label(frame, text='Dirección: ').grid(row = 3, column = 0)
		self.address = Entry(frame)
		self.address.grid(row = 3, column = 1)

		#City Input
		Label(frame, text='Ciudad: ').grid(row = 4, column = 0)
		defaultCity = StringVar(frame, value = 'Manizales')
		self.city = Entry(frame, textvariable = defaultCity)
		self.city.grid(row = 4, column = 1)

		#Phone Input
		Label(frame, text='Teléfono: ').grid(row = 5, column = 0)
		self.phone = Entry(frame)
		self.phone.grid(row = 5, column = 1)

		#Email Input
		Label(frame, text='Email: ').grid(row = 6, column = 0)
		self.email = Entry(frame)
		self.email.grid(row = 6, column = 1)

		#Button add client
		ttk.Button(frame, text='Agregar Cliente', command = self.addClient).grid(row = 7, columnspan = 2, sticky = W + E)

		#Output Messsages
		self.message = Label(text = '', fg = 'green')
		self.message.grid(row = 7, column = 0, columnspan = 2, sticky = W + E)

		#Table
		self.tree = ttk.Treeview(height = 10, columns = ('#0','#1','#2'))
		self.tree.grid(row = 9, column = 0, columnspan = 2)
		self.tree.heading('#0', text = 'Nombre', anchor = CENTER)
		self.tree.heading('#1', text = 'Representante', anchor = CENTER)
		self.tree.heading('#2', text = 'Telefono', anchor = CENTER)
		self.tree.heading('#3', text = 'Email', anchor = CENTER)

		#Button Editar y Eliminar
		ttk.Button(text = 'EDITAR', command = self.editClient).grid(row = 8, column = 0)
		ttk.Button(text = 'ELIMINAR', command = self.deleteClient).grid(row = 8, column = 1)
		ttk.Button(text = 'NUEVA COTIZACION', command = self.createCot).grid(row = 8, column = 2)

		#Llenando las filas de la tabla
		self.getClients()

	def getClients(self, parameters = ()):
		db = TinyDB('db/db.json')
		clients = db.all()

		#Limpia los datos existentes en la tabla
		records = self.tree.get_children()
		for element in records:
			self.tree.delete(element)

		#Obtiene los datos de la BD
		for client in clients:
			self.tree.insert('', 0, text = client["Client"], values = (client["Name"], client["Phone"], client["Email"]))


	def validate(self):
		return len(self.client.get()) != 0 and len(self.representative.get()) != 0 and len(self.address.get()) != 0 and len(self.city.get()) != 0 and len(self.phone.get()) > 6 and len(self.email.get()) != 0

	def addClient(self):
		if self.validate():
			db = TinyDB('db/db.json')
			Cliente = Query()

			client_name = self.client.get()
			client_name = client_name.upper()

			db.insert({'Client': client_name, 'Count': 0, 'Name' : self.representative.get(), 'Address' : self.address.get(), 'City' : self.city.get(), 'Phone' : self.phone.get(), 'Email' : self.email.get()})
			self.message['fg'] = 'green'
			self.message['text'] = "Cliente {} añadido exitosamente".format(client_name)

			#Crea carpeta que almacenara las cotizaciones
			os.system('mkdir "..\\%s"' % client_name)

			#Limpia el formulario 
			self.client.delete(0, END)
			self.representative.delete(0, END)
			self.address.delete(0, END)
			self.phone.delete(0, END)
			self.email.delete(0, END)
		else:
			self.message['fg'] = 'red'
			self.message['text'] = "Todos los datos son obligatorios"

		self.getClients()

	def deleteClient(self):
		db = TinyDB('db/db.json')
		Cliente = Query()
		
		try:
			self.tree.item(self.tree.selection())['text'][0]
		except IndexError as e:
			self.message['fg'] = 'red'
			self.message['text'] = 'Para eliminar debes seleccionar un cliente'
			return

		name = self.tree.item(self.tree.selection())['text']
		db.remove(Cliente.Client == name)
		self.message['fg'] = 'red'
		self.message['text'] = 'Cliente {} eliminado exitosamente'.format(name)
		self.getClients()

	def editClient(self):
		pass

	def createCot(self):
		db = TinyDB('db/db.json')
		Cliente = Query()
		
		try:
			self.tree.item(self.tree.selection())['text'][0]
		except IndexError as e:
			self.message['fg'] = 'red'
			self.message['text'] = 'Para crear una cotizacion debes seleccionar un cliente'
			return

		name = self.tree.item(self.tree.selection())['text']

		main.CrearCotizacion(name)

		self.message['fg'] = 'green'
		self.message['text'] = 'Cotizacion creada existosamente para {}'.format(name)


if __name__ == '__main__':
	window = Tk()
	app = Cliente(window)
	window.mainloop()
