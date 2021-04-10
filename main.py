import openpyxl
import numpy as nmp
wob1 = openpyxl.load_workbook('trabajo1.xlsx')

#Accedemos a las hojas de cálculo
num1 = wob1['Hoja1']

#Accedemos a las celdas y las transformamos en valores
A1 = num1['A1'].value
B1 = num1['B1'].value
C1 = num1['C1'].value
A3 = num1['A3'].value
B3 = num1['B3'].value
C3 = num1['C3'].value

print(A1, B1, C1, A3, B3, C3, "\n")

#Creamos el archivo y asignamos una hoja de calculo y guardamos el archivo
sol1 = openpyxl.Workbook()
hoja = sol1.active
sol1.save('sol1.xlsx')

#podemos realiza operaciónes por ejemplo multiplicando A1 por pi
numpi = A1 * nmp.pi
print("Multiplicamos A1 por el número pi: \n", numpi, "\n")

#tambien podemos hacer  una operación para encontrar las raices de un polinomio formado por C1,A1 y C3
numraiz = nmp.roots([C1, A1, C3])
print(" las raíces del polinomio: \n", numraiz, "\n")

#Realizamos una operación para establecer un rango en el cual se aumente un número tantas veces
ran = nmp.arange(A1, C1, A3)
print("El rango: \n", ran, "\n")

#tambien podemos crear un número complejo
numcom = complex(B3, C1)
print("Número complejo es: \n", numcom, "\n")

#Anexamos la información de las operaciones a la hoja que creamos en celdas específicas
hoja['E1'] = ("Multiplicamos A1 por el número pi: ")
hoja['E2'] = numpi
hoja['F1'] = ("las raíces del polinomio: ")
hoja['F2'] = str(numraiz)
hoja['G1'] = ("Obtener el rango: ")
hoja['G2'] = str(ran)
hoja['H1'] = ("número complejo es: ")
hoja['H2'] = str(numcom)

#Guardamos el archivo con las respuestas a los ejercicios
sol1.save('sol1.xlsx')