#Autor:    Janchooon

###################################################################

print('POR FAVOR COPIE EN EL CMD(ASUMIENDO QUE SE UTILIZA W10, SINO REALICE LA INSTALACION CORRESPONDIENTE EN SU SISTEMA OPERATIVO)\nCOMO ADMINISTRADOR LO SIGUIENTE:')
print('pip install pandas')
print('pip install xlrd')
print('pip install openpyxl')
print('SI YA LOS TIENE INSTALADOS OMITA ESTE MENSAJE\n')


##################################################################
# primero instalar esto a traves del cmd como administrador
# primero instalar esto:
#pip install pandas
#pip install xlrd

###################################################################

import pandas as pd

###################################################################

# Metodo para ordenar los datos de menor a mayor
def Ordenar(lista):
    # si el largo de la lista es 1 o 0 automaticamente estará ordenada
    if len(lista)>1:
        # Este proceso es O(log n) debido a que se realiza recursivamente y divide constsnemtemente las listas en 2
        # calcula el valor medio de la lista en numeros enteros
        medio = len(lista)//2
        # separa la lista en dos mitades
        derecha = lista[0:medio]
        izquierda = lista[medio:len(lista)]

        # recursivamente vuelve a separar los extremos de la lista
        Ordenar(derecha)
        Ordenar(izquierda)

        # El siguiente proceso se repite n cantidades de veces donde n es el largo de la lista O(n)
        
        #posiciones
        i=0
        j=0
        k=0
        
        # proceso de ordenamiento en donde normalente el largo de la lista se fragmenta hasta llegar a 1 y luego comparar sus valores
        while i < len(derecha) and j < len(izquierda):
            if derecha[i] <= izquierda[j]:
                lista[k]=derecha[i]
                i=i+1
            else:
                lista[k]=izquierda[j]
                j=j+1
            k=k+1

        while i < len(derecha):
            lista[k]=derecha[i]
            i=i+1
            k=k+1

        while j < len(izquierda):
            lista[k]=izquierda[j]
            j=j+1
            k=k+1

###################################################################

# Metodo de busqueda binaria
def Buscar (Lista, izq, der, dato):   
    # Caso base 
    if der >= izq: 

        #Calcula la mitad de la lista
        medio = izq + (der - izq)//2
       
  
        # Ve si el dato esta en medio de la lista
        if Lista[medio] == Lista[dato]: 
            return Lista[medio]
          
        # si el dato es menor se va hacia la izquierda y sigue así recursivamente hasta encontrarlo
        elif Lista[medio] > Lista[dato]: 
            return Buscar(Lista, izq, medio - 1, dato)
  
        # si el dato es menor se va hacia la izquierda y sigue así recursivamente hasta encontrarlo 
        else: 
            return Buscar(Lista, medio + 1, der, dato)
  
    else:
        return 'El alumno no existe'

###################################################################
#Menu
def Menu():
    print('N° |        NOMBRE')
    print('=============================')
    print(NO.Alumno)
    print('=============================')

    Condicion = True
    while Condicion:
        try:
            Alumnos = int(input('\nSeleccione a un alumno según su indice numerico: '))
            print('################################################')
            print('\nEl alumno seleccionado es',LNO[Alumnos][0],'\n')
            print('################################################')
            if int(Alumnos) > len(LNO)-1:
                print('\nEl alumno no existe, vuelve a intentar')
            else:
                Condicion = False
        except ValueError:
            print('\nEl alumno no existe, vuelve a intentar')

    Condicion2 = True
    while Condicion2:
        try:
            print('           NOTAS')
            print('===========================')
            print('         CONTROLES')
            print('===========================')
            print('1.-  C1')
            print('2.-  C2')
            print('3.-  C3')
            print('4.-  C4')
            print('===========================')
            print('       LABORATORIOS')
            print('===========================')
            print('5.-  L1')
            print('6.-  L2')
            print('7.-  L3')
            print('8.-  L4')
            print('9.-  L5')
            print('===========================')
            print('          EXAMEN')
            print('===========================')
            print('10.- EX')
            print('===========================')
            print('        PROMEDIOS')
            print('===========================')
            print('11.- PC')
            print('12.- PL')
            print('13.- PG')
            print('===========================')
            print('14.- EXPORTAR PLANILLA')
            print('15.- VOLVER')
            print('16.- SALIR')
            print('===========================')
            print(' * RESPONDA CON EL INDICE')
            print(' NUMERICO DE LA OPCION QUE')
            print('    DESEA SELECCIONAR     *')
            print('===========================')
            
            Opciones = int(input('Respuesta: '))
            if Opciones <= 16:
                if Opciones == 1:
                    print('\nLa nota de',LNO[Alumnos][0],'en el control 1 es',LNO[Alumnos][1],'\n')
                        
                elif Opciones == 2:
                    print('\nLa nota de',LNO[Alumnos][0],'en el control 2 es',LNO[Alumnos][2],'\n')
                        
                elif Opciones == 3:
                    print('\nLa nota de',LNO[Alumnos][0],'en el control 3 es',LNO[Alumnos][3],'\n')
                        
                elif Opciones == 4:
                    print('\nLa nota de',LNO[Alumnos][0],'en el control 4 es',LNO[Alumnos][4],'\n')

                elif Opciones == 5:
                    print('\nLa nota de',LNO[Alumnos][0],'en el laboratorio 1 es',LNO[Alumnos][5],'\n')
                        
                elif Opciones == 6:
                    print('\nLa nota de',LNO[Alumnos][0],'en el laboratorio 2 es',LNO[Alumnos][6],'\n')
                        
                elif Opciones == 7:
                    print('\nLa nota de',LNO[Alumnos][0],'en el laboratorio 3 es',LNO[Alumnos][7],'\n')
                        
                elif Opciones == 8:
                    print('\nLa nota de',LNO[Alumnos][0],'en el laboratorio 4 es',LNO[Alumnos][8],'\n')

                elif Opciones == 9:
                    print('\nLa nota de',LNO[Alumnos][0],'en el laboratorio 5 es',LNO[Alumnos][9],'\n')

                elif Opciones == 10:
                    print('\nLa nota de',LNO[Alumnos][0],'en el examen es',LNO[Alumnos][10],'\n')
                    
                elif Opciones == 11:
                    print('\nEl promedio de controles de',LNO[Alumnos][0],'es',LNO[Alumnos][11],'\n')
                        
                elif Opciones == 12:
                    print('\nEl promedio de laboratorios de',LNO[Alumnos][0],'es',LNO[Alumnos][12],'\n')
                        
                elif Opciones == 13:
                    print('\nEl promedio general de',LNO[Alumnos][0],'es',LNO[Alumnos][13],'\n')

                elif Opciones == 14:
                    Exportar = 'Notas_Ordenadas.xlsx'
                    print(NO.to_excel(Exportar))
                    print('############################')
                    print('LA PLANILLA A SIDO EXPORTADA')
                    print('############################\n')
                    return Menu()

                elif Opciones == 15:
                    return Menu()

                elif Opciones == 16:
                    Condicion2 = False
                    print('HASTA LA PROXIMA')
                        
                else:
                    print('\nEsa opcion no está disponible, vuelve a intentar\n')
            else:
                print('\nEsa opcion no está disponible, vuelve a intentar\n')
                    
        except ValueError:
            print('\nEsa opcion no está disponible, vuelve a intentar\n')


###################################################################

# Exportar datos del excel
Excel = pd.ExcelFile('Planilla_Notas.xlsx')
Notas = Excel.parse('Notas')

# Pasa todas las iniciales de apellido y nombre a mayuscula para ordenar bien
Colu = []
for i in range(0,len(Notas)):
    Nom = Notas.Alumno[i]
    Nom = Nom.title()
    Colu.append(Nom)

Colu = pd.DataFrame(Colu)

# Se agregan los nombres ya arreglados al DataFrame
Notas['Alumno'] = Colu
    

# Lista para agregar filas de DataFrame como listas
ListaN = []

# Separa las filas, las transforma en listas y las agrega a ListaN
for indice_fila, fila in Notas.iterrows():
    ListaN.append(list(fila))

   

# Llamar a la funcion orden para ordenasr las listas dentro de listas
Ordenar(ListaN)

# Listas para agregar los promedios
PC = []
PL = []
PG = []

#Promedio de controles
for i in ListaN:
    SumaC = 0
    cant = 0
    for j in i[1:5]:
        SumaC = SumaC + j
    PromC = SumaC/4
    PC.append(PromC)

#Promedio de laboratorios
for i in ListaN:
    SumaL = 0
    for j in i[5:10]:
        SumaL = SumaL + j
    PromL = SumaL/5
    PL.append(PromL)

#Promedio General
cont = 0
for i in range(0,len(ListaN)):
    cont += 1
    Con = PC[i]
    Lab = PL[i]
    Exa = (ListaN[i][10])
    PromG = (Con*0.4) + (Lab*0.5) + (Exa*0.1)
    PG.append(PromG)


# DataFrame con las notas ordenadas y con los promedios agregados
NotasOrdenadas = pd.DataFrame(ListaN,columns = ['Alumno','C1','C2','C3','C4','L1','L2','L3','L4','L5','EX'])
NotasOrdenadas['PC'] = PC
NotasOrdenadas['PL'] = PL
NotasOrdenadas['PG'] = PG

# Redondeo de las notas
NO = round(NotasOrdenadas,1)


# Lista para agregar filas de DataFrame como listas
LNO = []

# Separa las dilas, las transforma en listas y las agrega a ListaN
for indiceOrdeado, Fordenada in NO.iterrows():
    LNO.append(list(Fordenada))

# Ejecuta el programa
print(Menu())
