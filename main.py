from openpyxl import Workbook
from datetime import datetime

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


if __name__ == "__main__":
    app = App()
