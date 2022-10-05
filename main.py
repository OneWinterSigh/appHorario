from openpyxl import Workbook
from datetime import datetime
import tkinter as tk
from tkinter import ttk


class PersonaHorario:
    def __init__(self, nombre, codigo):
        posiblescodigos = ['01','02','03','04','05','06','07','08','09','10']
        self.codigo = codigo
        
    def getCodigo(self):
        return self.codigo

    def getCodigoTrabajador(self, trabajador)
             

class Almacen:
    def __init__(self):
        self.almacenDePersonas = []
        self.almacenDeNombres = []
        self.almacenHorasAcumuladas = []

    def getAlmacenDePersonas(self):
        return self.almacenDePersonas

    def getAlmacenDeNombres(self):
        return self.almacenDeNombres

    def getAlmacenHorasAcumuladas(self):
        return self.almacenHorasAcumuladas


class Funcionando:
    def __init__(self):
        self.i = 1
        self.condicion = True
        self.almacenDeTodo = Almacen()

    def getAlmacenDeTodo(self):
        return self.almacenDeTodo

    def pideLosDatos(self):
        while self.i > 0 and self.condicion:
            userInput = input("Quien es: ")

            # Para el programa
            if userInput == "p":
                break

            # Almacena los nombres en un array a parte
            if userInput not in self.almacenDeTodo.almacenDePersonas:
                self.almacenDeTodo.almacenDeNombres.append(userInput)

            userInputEntrada = input("¿Hora entrada? ")
            userInputSalida = input("¿Hora salida? ")

            variableProvisional = PersonaHorario(userInput, userInputEntrada, userInputSalida)
            variableProvisional.calculoDeHoras()

            self.almacenDeTodo.almacenDePersonas.append(variableProvisional.getFila())

    def calcularHorasDeCadaPersona(self):
        i = 0
        j = 0
        contador = 0
        cuenta = 0
        while i < len(self.almacenDeTodo.almacenDeNombres):
            while j < len(self.almacenDeTodo.almacenDePersonas):
                if self.almacenDeTodo.almacenDeNombres[i] == self.almacenDeTodo.almacenDePersonas[j]:
                    cuenta = cuenta + int(dir(self.almacenDeTodo.almacenDePersonas[j]).getHorasDeEseDia())
                else:
                    j = j + 1
            contador = contador + 1
            self.almacenDeTodo.almacenHorasAcumuladas.append(cuenta)
            cuenta = 0
            i = i + 1

    def pintarExcel(self):
        wb = Workbook()
        dest_filename = 'prueba12.xlsx'

        ws1 = wb.active
        ws1.title = "range names"

        i = 0
        j = 0
        while j < len(self.almacenDeTodo.almacenDePersonas):
            ws1.append(self.almacenDeTodo.almacenDePersonas[j])
            j = j + 1

        wb.save('/Users/germanruizcabello/projects/Python/appHorario/excels/' + dest_filename)

    def activate(self):
        self.pideLosDatos()

        pregunta = input("¿Necesitas algo más? (y/n)")
        if pregunta == "y":
            pregunta2 = input("¿Qué puedo hacer por ti? (puedo calcular las horas totales) (y/n)")
            if pregunta2 == "y":
                self.calcularHorasDeCadaPersona()
            else:
                pass
        else:
            pass

        print(self.almacenDeTodo.getAlmacenHorasAcumuladas())

        self.pintarExcel()

"""
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Tkinter StringVar')
        self.geometry("500x500+200+200")

        self.name_var = tk.StringVar()
        self.horaentrada_var = tk.StringVar()
        self.horasalida_var = tk.StringVar() 

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.columnconfigure(2, weight=1)

        self.create_widgets()

    def create_widgets(self):

        padding = {'padx': 5, 'pady': 5}
        # label
        ttk.Label(self, text='Introduce código:').grid(column=0, row=0, **padding)

        # Entry
        name_entry = ttk.Entry(self, textvariable=self.name_var)
        name_entry.grid(column=1, row=0, **padding)
        name_entry.focus()

        # Button
        submit_button = ttk.Button(self, text='Submit', command=self.submit)
        submit_button.grid(column=2, row=0, **padding)

        # Output label
        self.output_label = ttk.Label(self)
        self.output_label.grid(column=0, row=1, columnspan=3, **padding)

    def submit(self):
        self.output_label.config(text=self.name_var.get())

"""
if __name__ == "__main__":
    app = App()
    app.mainloop()