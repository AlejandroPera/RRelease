# -- coding: cp1252 --
"""Este programa hace uso de la librería pygame para crear una interfaz donde se despliega un videojuego dirigido para alumnos con
discapacidad intelectual a nivel primaria. El videjuego permite el aprendizaje y reforzamiento de ciertos campos semánticos relevantes
en la educación de los alumnos. Estos campos semanticos fueron desarrollados mediante el uso de clases y objetos, así como, otras
funciones tales como la creación de un nuevo usurio, estadisticas del alumno,la comuniación entre la botonera, etc."""

"""IMPORTANTE: Antes de correr el programa asegúrate de tener instaladas las siguiente librerías:

              pygame
              pandas
              openpyxl
              xlrd
              xlwritter
              matplotlib
              pillow
              
             de lo contrario usar el comando pip install (nombre de librería) en el intérprete de comandos 'Command prompt'

             Así como, haber cargado previamente el programa Botonera.ino que se encuentra en la carpeta Arduino en nuestro microcontrolador"""


import sys
import random
import serial #Librería que permite la comuniación con arduino
import matplotlib #Librería que nos permite graficar el progreso del alumno
import matplotlib.pyplot as plt
import numpy as np #Librería que brinda soporte para vectores y matrices
import pandas #Librería que permite escribir datos en excel
import pygame # Librería que permite la creación del videojuego
import time
from pygame.locals import *
from pandas import ExcelWriter
from PIL import Image # Librería que permite ajustar el tamaño de imagenes


pygame.init()

modes = pygame.display.list_modes(32)
screen = pygame.display.set_mode(modes[0],RESIZABLE,32)
pygame.display.set_caption('Videojuego')

screen.fill((255,255,255))
pygame.display.update()

# Fuentes a utilizar en pygame
fuente = pygame.font.SysFont('Verdana',36)
fuente2 = pygame.font.SysFont('Arial', 32)
fuente3 =pygame.font.Font(None,32)
fuente4 =pygame.font.Font(None,86)
fuente5 =pygame.font.Font(None,20)

# Música y sonidos a utilizar en pygame
pygame.mixer.music.load("Sonidos/intro.mp3")
pygame.mixer.music.play(1)
pygame.mixer.music.set_volume(0.1)

aplausos = pygame.mixer.Sound("Sonidos/aplausos.wav")
aplausos.set_volume(0.1)
intentalo = pygame.mixer.Sound("Sonidos/intentalo.wav")
intentalo.set_volume(0.5)

azul = pygame.mixer.Sound("Sonidos/azul.wav")
azul.set_volume(0.5)
amarillo = pygame.mixer.Sound("Sonidos/amarillo.wav")
amarillo.set_volume(0.5)
rosa = pygame.mixer.Sound("Sonidos/rosa.wav")
rosa.set_volume(0.5)
verde = pygame.mixer.Sound("Sonidos/verde.wav")
verde.set_volume(0.5)
feliz = pygame.mixer.Sound("Sonidos/feliz.wav")
feliz.set_volume(0.5)
triste = pygame.mixer.Sound("Sonidos/triste.wav")
triste.set_volume(0.5)
enojada = pygame.mixer.Sound("Sonidos/enojada.wav")
enojada.set_volume(0.5)
sorprendida = pygame.mixer.Sound("Sonidos/sorprendida.wav")
sorprendida.set_volume(0.5)
oso = pygame.mixer.Sound("Sonidos/oso.wav")
oso.set_volume(0.5)
perro = pygame.mixer.Sound("Sonidos/perro.wav")
perro.set_volume(0.5)
gato = pygame.mixer.Sound("Sonidos/gato.wav")
gato.set_volume(0.5)
buho = pygame.mixer.Sound("Sonidos/buho.wav")
buho.set_volume(0.5)

uno = pygame.mixer.Sound("Sonidos/uno.wav")
uno.set_volume(0.5)
dos = pygame.mixer.Sound("Sonidos/dos.wav")
dos.set_volume(0.5)
tres = pygame.mixer.Sound("Sonidos/tres.wav")
tres.set_volume(0.5)
cuatro = pygame.mixer.Sound("Sonidos/cuatro.wav")
cuatro.set_volume(0.5)
cinco = pygame.mixer.Sound("Sonidos/cinco.wav")
cinco.set_volume(0.5)
seis = pygame.mixer.Sound("Sonidos/seis.wav")
seis.set_volume(0.5)
siete = pygame.mixer.Sound("Sonidos/siete.wav")
siete.set_volume(0.5)
ocho = pygame.mixer.Sound("Sonidos/ocho.wav")
ocho.set_volume(0.5)
nueve = pygame.mixer.Sound("Sonidos/nueve.wav")
nueve.set_volume(0.5)
diez = pygame.mixer.Sound("Sonidos/diez.wav")
diez.set_volume(0.5)

# Lectura de nuestro puerto serie de arduino
PuertoSerie = serial.Serial('COM5',9600)


class Botonera(object):
    def boton(self):
        """Leé y convierte los valores que manda arduino
            Devuelve los número 1 al 4 dependiendo el boton presionado"""
        sArduino = PuertoSerie.readline()
        numero = list(sArduino)
        self.eleccion = int(numero[0])
        return self.eleccion

# Campo semántico Colores        
class Colores(object):

    def _init_(self):
        self.azul = pygame.image.load("Imagenes/Colores/azul.png")
        self.amarillo = pygame.image.load("Imagenes/Colores/amarillo.png")
        self.verde = pygame.image.load("Imagenes/Colores/verde.png")
        self.rosa = pygame.image.load("Imagenes/Colores/rosa.png")
        self.opciones = ['azul','amarillo','verde','rosa']

        self.indi = 0

    def  eleccion(self):
        """Genera 5 listas self.juego(Campo semántico), self.estatus(Correcto o incorrecto representado en 1 y 0),
            self.tcolores(tiempo de reacción de respuesta) con ciertos valores dependiendo si el alumno respondió correctamente o no y en cuánto tiempo lo hizo,
            self.real(El valor seleccionado del usuario), self.ideal(El valor correcto que se debió seleccionar)"""
        self.juego = []
        self.estatus = []
        self.real = []
        self.ideal = []
        self.tcolores = []

        salida = 0
        uwu = 0
        posiciones = [0,350,700,1050]
        pos = random.sample(posiciones,len(posiciones))
        self.indi = 1
        
        while salida == 0:

            marca = 0

            screen.fill((255,255,255))
            campo.mensajes() 
            fig = random.choice(color.opciones)
            mensaje = fuente.render('Selecciona el color ' + fig,1,(0,0,255),True)
            screen.blit(color.azul,(pos[0],200))
            screen.blit(color.amarillo,(pos[1],200))
            screen.blit(color.verde,(pos[2],200))
            screen.blit(color.rosa,(pos[3],180))
            pygame.display.update()

            # Delay para dar tiempo a comando de voz terminar de decir instrucción
            if  fig == 'azul'and salida == 0:
                azul.play()
                pygame.time.delay(1500)

            elif  fig == 'amarillo'and salida == 0:
                amarillo.play()
                pygame.time.delay(1500)

            elif  fig == 'verde'and salida == 0:
                verde.play()
                pygame.time.delay(1500)

            elif  fig == 'rosa'and salida == 0:
                rosa.play()
                pygame.time.delay(1500)

            screen.blit(mensaje,(460,40))
            pygame.display.update()

            if pos[0] == 0:
                azu = 1
            elif pos[0] == 350:
                azu = 2
            elif pos[0] == 700:
                azu = 3
            else:
                azu = 4

            if pos[1] == 0:
                ama = 1
            elif pos[1] == 350:
                ama = 2
            elif pos[1] == 700:
                ama = 3
            else:
                ama = 4

            if pos[2] == 0:
                ver = 1
            elif pos[2] == 350:
                ver = 2
            elif pos[2] == 700:
                ver = 3
            else:
                ver = 4

            if pos[3] == 0:
                ros = 1
            elif pos[3] == 350:
                ros = 2
            elif pos[3] == 700:
                ros = 3
            else:
                ros = 4


            while marca == 0:

                inicio = time.time()

                for event in pygame.event.get():
                    if event.type == QUIT: 
                        pygame.quit()
                        sys.exit()
                    elif event.type == KEYDOWN:
                        if event.key == K_s:
                            uwu = 1

                if fig == 'azul' and bot.boton() == azu:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tcolores.append(tiempo)

                    self.juego.append("Colores")
                    self.real.append("Azul")
                    self.estatus.append(1)
                    marca = 1                    
                elif fig == 'amarillo' and bot.boton() == ama:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tcolores.append(tiempo)
                    
                    self.juego.append("Colores")
                    self.real.append("Amarillo")
                    self.estatus.append(1)
                    marca = 1     
                elif fig == 'verde' and bot.boton() == ver:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tcolores.append(tiempo)
                    
                    self.juego.append("Colores")
                    self.real.append("Verde")
                    self.estatus.append(1)
                    marca = 1     
                elif fig == 'rosa' and bot.boton() == ros:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tcolores.append(tiempo)
                    
                    self.juego.append("Colores")
                    self.real.append("Rosa")
                    self.estatus.append(1)
                    marca = 1     
                else:
                    intentalo.play()
                    pygame.time.delay(1500)
                    self.tcolores.append("NA")

                    self.juego.append("Colores")                
                    self.ideal.append(fig)
                    self.estatus.append(0)

                if uwu == 1:
                    marca = 1
            if uwu == 1:
                salida = 1
        estadistica.escribir()
        self.indi = 0
        campo.seleccion()

    def respuestas(self):
        """Leé la lista self.rcolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            jueg = self.juego
            return jueg
        else:
            return []
        
    def respuestas2(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            rea = self.real
            return rea
        else:
            return []

    def respuestas3(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            idea = self.ideal
            return idea
        else:
            return []

    def imagenes(self):
        """Leé la lista self.icolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            estatu = self.estatus
            return estatu
        else:
            return []

    def tiempos(self):
        """Leé la lista self.tcolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            tiemp = self.tcolores
            return tiemp
        else:
            return []

# Campo semántico Emociones        
class Emociones(object):

    def _init_(self):
        self.feliz = pygame.image.load("Imagenes/Emociones/feliz.png")
        self.triste = pygame.image.load("Imagenes/Emociones/triste.png")
        self.enojada = pygame.image.load("Imagenes/Emociones/enojada.png")
        self.sorprendida = pygame.image.load("Imagenes/Emociones/sorprendida.png")
        self.opciones = ['feliz','triste','enojada','sorprendida']

        self.indi = 0

    def eleccion(self):
        """Genera 3 listas self.remociones(Respuestas correctas o incorrectas), self.iemociones(Imagen correcta o incorrecta),
            self.temociones(tiempo de reacción de respuesta) con ciertos valores dependiendo si el alumno respondió correctamente o no y en cuánto tiempo lo hizo"""
        self.juego = []
        self.estatus = []
        self.real = []
        self.ideal = []
        self.temociones = []

        salida = 0
        uwu = 0
        posiciones = [0,350,700,1050]
        pos = random.sample(posiciones,len(posiciones))
        self.indi = 1
        
        while salida == 0:

            marca = 0

            screen.fill((255,255,255))
            campo.mensajes() 
            fig = random.choice(emocion.opciones)
            mensaje = fuente.render('Selecciona la carita ' + fig,1,(0,0,255),True)       
            screen.blit(emocion.feliz,(pos[0],200))
            screen.blit(emocion.triste,(pos[1],200))
            screen.blit(emocion.enojada,(pos[2],200))
            screen.blit(emocion.sorprendida,(pos[3],200))
            pygame.display.update()

            # Delay para dar tiempo a comando de voz terminar de decir instrucción
            if  fig == 'feliz'and salida == 0:
                feliz.play()
                pygame.time.delay(1500)

            elif  fig == 'triste'and salida == 0:
                triste.play()
                pygame.time.delay(1500)

            elif  fig == 'enojada'and salida == 0:
                enojada.play()
                pygame.time.delay(1500)

            elif  fig == 'sorprendida'and salida == 0:
                sorprendida.play()
                pygame.time.delay(1500)

            screen.blit(mensaje,(445,40)) 
            pygame.display.update()

            if pos[0] == 0:
                feli = 1
            elif pos[0] == 350:
                feli = 2
            elif pos[0] == 700:
                feli = 3
            else:
                feli = 4

            if pos[1] == 0:
                tris = 1
            elif pos[1] == 350:
                tris = 2
            elif pos[1] == 700:
                tris = 3
            else:
                tris = 4

            if pos[2] == 0:
                enoj = 1
            elif pos[2] == 350:
                enoj = 2
            elif pos[2] == 700:
                enoj = 3
            else:
                enoj = 4

            if pos[3] == 0:
                sorp = 1
            elif pos[3] == 350:
                sorp = 2
            elif pos[3] == 700:
                sorp = 3
            else:
                sorp = 4
                

            while marca == 0:

                inicio = time.time()

                for event in pygame.event.get():
                    if event.type == QUIT: 
                        pygame.quit()
                        sys.exit()
                    elif event.type == KEYDOWN:
                        if event.key == K_s:
                            uwu = 1

                if fig == 'feliz' and bot.boton() == feli:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.temociones.append(tiempo)

                    self.juego.append("Emociones")
                    self.real.append("Feliz")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'triste' and bot.boton() == tris:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.temociones.append(tiempo)

                    self.juego.append("Emociones")
                    self.real.append("Triste")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'enojada' and bot.boton() == enoj:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.temociones.append(tiempo)

                    self.juego.append("Emociones")
                    self.real.append("Enojada")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'sorprendida' and bot.boton() == sorp:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.temociones.append(tiempo)

                    self.juego.append("Emociones")
                    self.real.append("Sorpresa")
                    self.estatus.append(1)
                    marca = 1
                else:
                    intentalo.play()
                    pygame.time.delay(1500)
                    self.temociones.append("NA")

                    self.juego.append("Emociones")                
                    self.ideal.append(fig)
                    self.estatus.append(0)

                if uwu == 1:
                    marca = 1
            if uwu == 1:
                salida = 1
        estadistica.escribir()
        self.indi = 0
        campo.seleccion()

    def respuestas(self):
        """Leé la lista self.rcolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            jueg = self.juego
            return jueg
        else:
            return []
        
    def respuestas2(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            rea = self.real
            return rea
        else:
            return []

    def respuestas3(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            idea = self.ideal
            return idea
        else:
            return []

    def imagenes(self):
        """Leé la lista self.icolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            estatu = self.estatus
            return estatu
        else:
            return []

    def tiempos(self):
        """Leé la lista self.temociones
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Emociones, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            tiemp = self.temociones
            return tiemp
        else:
            return []

# Campo semántico Animales
class Animales(object):

    def _init_(self):
        self.perro = pygame.image.load("Imagenes/Animales/perro.png")
        self.gato = pygame.image.load("Imagenes/Animales/gato.png")
        self.buho = pygame.image.load("Imagenes/Animales/buho.png")
        self.oso = pygame.image.load("Imagenes/Animales/oso.png")
        self.opciones = ['perro','gato','buho','oso']

        self.indi = 0

    def eleccion(self):
        """Genera 3 listas self.ranimales(Respuestas correctas o incorrectas), self.ianimales(Imagen correcta o incorrecta),
            self.tanimaless(tiempo de reacción de respuesta) con ciertos valores dependiendo si el alumno respondió correctamente o no y en cuánto tiempo lo hizo"""
        self.juego = []
        self.estatus = []
        self.real = []
        self.ideal = []
        self.tanimales = []

        salida = 0
        uwu = 0
        posiciones = [0,350,700,1050]
        pos = random.sample(posiciones,len(posiciones))
        self.indi = 1
        
        while salida == 0:

            marca = 0

            screen.fill((255,255,255))
            campo.mensajes() 
            fig = random.choice(animal.opciones)
            mensaje = fuente.render('Selecciona al ' + fig,1,(0,0,255),True)
            screen.blit(animal.perro,(pos[0],200))
            screen.blit(animal.gato,(pos[1],200))
            screen.blit(animal.buho,(pos[2],160))
            screen.blit(animal.oso,(pos[3],170))

            pygame.display.update()

            # Delay para dar tiempo a comando de voz terminar de decir instrucción
            if  fig == 'perro'and salida == 0:
                perro.play()
                pygame.time.delay(1500)

            elif  fig == 'gato'and salida == 0:
                gato.play()
                pygame.time.delay(1500)

            elif  fig == 'oso'and salida == 0:
                oso.play()
                pygame.time.delay(1500)

            elif  fig == 'buho'and salida == 0:
                buho.play()
                pygame.time.delay(1500)

            screen.blit(mensaje,(500,40)) 
            pygame.display.update()


            if pos[0] == 0:
                per = 1
            elif pos[0] == 350:
                per = 2
            elif pos[0] == 700:
                per = 3
            else:
                per = 4

            if pos[1] == 0:
                gat = 1
            elif pos[1] == 350:
                gat = 2
            elif pos[1] == 700:
                gat = 3
            else:
                gat = 4

            if pos[2] == 0:
                buh = 1
            elif pos[2] == 350:
                buh = 2
            elif pos[2] == 700:
                buh = 3
            else:
                buh = 4

            if pos[3] == 0:
                os = 1
            elif pos[3] == 350:
                os = 2
            elif pos[3] == 700:
                os = 3
            else:
                os = 4
                

            while marca == 0:

                inicio = time.time()

                for event in pygame.event.get():
                    if event.type == QUIT: 
                        pygame.quit()
                        sys.exit()
                    elif event.type == KEYDOWN:
                        if event.key == K_s:
                            uwu = 1

                if fig == 'perro' and bot.boton() == per:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tanimales.append(tiempo)

                    self.juego.append("Animales")
                    self.real.append("Perro")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'gato' and bot.boton() == gat:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tanimales.append(tiempo)

                    self.juego.append("Animales")
                    self.real.append("Gato")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'buho' and bot.boton() == buh:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tanimales.append(tiempo)
                    
                    self.juego.append("Animales")
                    self.real.append("Buho")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'oso' and bot.boton() == os:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tanimales.append(tiempo)

                    self.juego.append("Animales")
                    self.real.append("Oso")
                    self.estatus.append(1)
                    marca = 1
                else:
                    intentalo.play()
                    pygame.time.delay(1500)
                    self.tanimales.append("NA")

                    self.juego.append("Animales")                
                    self.ideal.append(fig)
                    self.estatus.append(0)
                    
                if uwu == 1:
                    marca = 1
            if uwu == 1:
                salida = 1
        estadistica.escribir()
        self.indi = 0
        campo.seleccion()                            

    def respuestas(self):
        """Leé la lista self.rcolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            jueg = self.juego
            return jueg
        else:
            return []
        
    def respuestas2(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            rea = self.real
            return rea
        else:
            return []

    def respuestas3(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            idea = self.ideal
            return idea
        else:
            return []

    def imagenes(self):
        """Leé la lista self.icolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            estatu = self.estatus
            return estatu
        else:
            return []

    def tiempos(self):
        """Leé la lista self.tanimales
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Animales, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            tiemp = self.tanimales
            return tiemp
        else:
            return []

# Campo semántico Números
class Numeros(object):  

    def _init_(self):
        self.uno = pygame.image.load("Imagenes/Numeros/uno.png")
        self.dos = pygame.image.load("Imagenes/Numeros/dos.png")
        self.tres = pygame.image.load("Imagenes/Numeros/tres.png")
        self.cuatro = pygame.image.load("Imagenes/Numeros/cuatro.png")
        self.cinco = pygame.image.load("Imagenes/Numeros/cinco.png")
        self.seis = pygame.image.load("Imagenes/Numeros/seis.png")
        self.siete = pygame.image.load("Imagenes/Numeros/siete.png")
        self.ocho = pygame.image.load("Imagenes/Numeros/ocho.png")
        self.nueve = pygame.image.load("Imagenes/Numeros/nueve.png")
        self.diez = pygame.image.load("Imagenes/Numeros/diez.png")
        
        self.opciones = ['uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez']
        self.indi = 0

    def eleccion(self):
        """Genera 3 listas self.rnumeros(Respuestas correctas o incorrectas), self.inumeros(Imagen correcta o incorrecta),
            self.tnumeros(tiempo de reacción de respuesta) con ciertos valores dependiendo si el alumno respondió correctamente o no y en cuánto tiempo lo hizo"""
        self.juego = []
        self.estatus = []
        self.real = []
        self.ideal = []
        self.tnumeros = []
        self.monedas = []

        salida = 0
        uwu = 0
        opciones = ['uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez']
        mon = random.sample(opciones,4)

        for i in mon:
            if i == 'uno':
                self.monedas.append(self.uno)
            elif i == 'dos':
                self.monedas.append(self.dos)            
            elif i == 'tres':
                self.monedas.append(self.tres)
            elif i == 'cuatro':
                self.monedas.append(self.cuatro)  
            elif i == 'cinco':
                self.monedas.append(self.cinco)
            elif i == 'seis':
                self.monedas.append(self.seis)            
            elif i == 'siete':
                self.monedas.append(self.siete)
            elif i == 'ocho':
                self.monedas.append(self.ocho)
            elif i == 'nueve':
                self.monedas.append(self.nueve)
            else:
                self.monedas.append(self.diez)  
        
        self.indi = 1

        while salida == 0:
                    
            marca = 0
   
            screen.fill((255,255,255))
            campo.mensajes()
            fig = random.choice(mon)
            if fig == 'uno':
                mensaje = fuente.render('Selecciona la moneda',1,(0,0,255),True)
            else:
                mensaje = fuente.render('Selecciona las ' + fig + ' monedas',1,(0,0,255),True)
            screen.blit(self.monedas[0],(40,200))
            screen.blit(self.monedas[1],(390,200))
            screen.blit(self.monedas[2],(740,200))
            screen.blit(self.monedas[3],(1090,200))
            pygame.display.update()

            # Delay para dar tiempo a comando de voz terminar de decir instrucción
            if  fig == 'uno'and salida == 0:
                uno.play()
                pygame.time.delay(1500)

            elif  fig == 'dos'and salida == 0:
                dos.play()
                pygame.time.delay(1500)

            elif  fig == 'tres'and salida == 0:
                tres.play()
                pygame.time.delay(1500)

            elif  fig == 'cuatro'and salida == 0:
                cuatro.play()
                pygame.time.delay(1500)

            elif  fig == 'cinco'and salida == 0:
                cinco.play()
                pygame.time.delay(1500)

            elif  fig == 'seis'and salida == 0:
                seis.play()
                pygame.time.delay(1500)

            elif  fig == 'siete'and salida == 0:
                siete.play()
                pygame.time.delay(1500)

            elif  fig == 'ocho'and salida == 0:
                ocho.play()
                pygame.time.delay(1500)

            elif  fig == 'nueve'and salida == 0:
                nueve.play()
                pygame.time.delay(1500)

            elif  fig == 'diez'and salida == 0:
                diez.play()
                pygame.time.delay(1500)

            screen.blit(mensaje,(450,40))
            pygame.display.update()

            if mon[0] == 'uno':
                un = 1
            elif mon[1] == 'uno':
                un = 2
            elif mon[2] == 'uno':
                un = 3
            elif mon[3] == 'uno':
                un = 4

            if mon[0] == 'dos':
                do = 1
            elif mon[1] == 'dos':
                do = 2
            elif mon[2] == 'dos':
                do = 3
            elif mon[3] == 'dos':
                do = 4

            if mon[0] == 'tres':
                tre = 1
            elif mon[1] == 'tres':
                tre = 2
            elif mon[2] == 'tres':
                tre = 3
            elif mon[3] == 'tres':
                tre = 4

            if mon[0] == 'cuatro':
                cuat = 1
            elif mon[1] == 'cuatro':
                cuat = 2
            elif mon[2] == 'cuatro':
                cuat = 3
            elif mon[3] == 'cuatro':
                cuat = 4

            if mon[0] == 'cinco':
                cinc = 1
            elif mon[1] == 'cinco':
                cinc = 2
            elif mon[2] == 'cinco':
                cinc = 3
            elif mon[3] == 'cinco':
                cinc = 4

            if mon[0] == 'seis':
                sei = 1
            elif mon[1] == 'seis':
                sei = 2
            elif mon[2] == 'seis':
                sei = 3
            elif mon[3] == 'seis':
                sei = 4

            if mon[0] == 'siete':
                sie = 1
            elif mon[1] == 'siete':
                sie = 2
            elif mon[2] == 'siete':
                sie = 3
            elif mon[3] == 'siete':
                sie = 4

            if mon[0] == 'ocho':
                och = 1
            elif mon[1] == 'ocho':
                och = 2
            elif mon[2] == 'ocho':
                och = 3
            elif mon[3] == 'ocho':
                och = 4

            if mon[0] == 'nueve':
                nue = 1
            elif mon[1] == 'nueve':
                nue = 2
            elif mon[2] == 'nueve':
                nue = 3
            elif mon[3] == 'nueve':
                nue = 4

            if mon[0] == 'diez':
                die = 1
            elif mon[1] == 'diez':
                die = 2
            elif mon[2] == 'diez':
                die = 3
            elif mon[3] == 'diez':
                die = 4
                         
            while marca == 0:

                inicio = time.time()

                for event in pygame.event.get():
                    if event.type == QUIT: 
                        pygame.quit()
                        sys.exit()
                    elif event.type == KEYDOWN:
                        if event.key == K_s:
                            uwu = 1

                if  fig == 'uno' and bot.boton() == un:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Uno")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'dos' and bot.boton() == do:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Dos")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'tres' and bot.boton() == tre:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Tres")
                    self.estatus.append(1)
                    marca = 1
                elif fig == 'cuatro' and bot.boton() == cuat:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Cuatro")
                    self.estatus.append(1)
                    marca = 1

                elif fig == 'cinco' and bot.boton() == cinc:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Cinco")
                    self.estatus.append(1)
                    marca = 1

                elif fig == 'seis' and bot.boton() == sei:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Seis")
                    self.estatus.append(1)
                    marca = 1

                elif fig == 'siete' and bot.boton() == sie:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Siete")
                    self.estatus.append(1)
                    marca = 1
                    
                elif fig == 'ocho' and bot.boton() == och:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Ocho")
                    self.estatus.append(1)
                    marca = 1

                elif fig == 'nueve' and bot.boton() == nue:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Nueve")
                    self.estatus.append(1)
                    marca = 1
                    
                elif fig == 'diez' and bot.boton() == die:
                    aplausos.play()
                    pygame.time.delay(1500)

                    final= time.time ()
                    tiempo = round(final-inicio,3)
                    self.tnumeros.append(tiempo)

                    self.juego.append("Numeros")
                    self.real.append("Diez")
                    self.estatus.append(1)
                    marca = 1
                
                else:
                    intentalo.play()
                    pygame.time.delay(1500)
                    self.tnumeros.append("NA")

                    self.juego.append("Numeros")                
                    self.ideal.append(fig)
                    self.estatus.append(0)

                if uwu == 1:
                    marca = 1
            if uwu == 1:
                salida = 1
        estadistica.escribir()
        self.indi = 0
        campo.seleccion()

    def respuestas(self):
        """Leé la lista self.rcolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            jueg = self.juego
            return jueg
        else:
            return []
        
    def respuestas2(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            rea = self.real
            return rea
        else:
            return []

    def respuestas3(self):
        """ Leé la lista self.real
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario rergesa una lista vacia"""
        if self.indi == 1:
            idea = self.ideal
            return idea
        else:
            return []

    def imagenes(self):
        """Leé la lista self.icolores 
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Colores, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            estatu = self.estatus
            return estatu
        else:
            return []

    def tiempos(self):
        """Leé la lista self.tnumeros
            Devuelve esa lista dependiendo si se entró a trabajar en el campo semántico Números, de lo contrario regresa una lista vacia"""
        if self.indi == 1:
            tiemp = self.tnumeros
            return tiemp
        else:
            return []


class Portada(object):

    def _init_(self):
        self.portada = pygame.image.load("Imagenes/maquina1.png")

    def cargando(self):
       
        """Crea una barra que simula que el juego está cargando"""    
        barPos      = (screen.get_width()/2-100,screen.get_height()/2-10)
        barSize     = (200, 20)
        borderColor = (0, 0, 0)
        barColor    = (0, 128, 0)

        cargando = fuente5.render('Cargando...',1,(0,0,0),True)

        max_a = int(378/2)
        
        for i in range(max_a):
            screen.fill((255,255,255))
            screen.blit(cargando,((screen.get_width()/2)-(cargando.get_width()/2),(screen.get_height()/2)-28))
            pygame.draw.rect(screen, borderColor, (barPos, barSize), 1)
            innerPos  = (barPos[0]+3, barPos[1]+3)
            innerSize = ((barSize[0]-194+i), barSize[1]-6)
            pygame.draw.rect(screen, barColor, (innerPos, innerSize))
            pygame.display.update()
            pygame.time.wait(5)

            for event in pygame.event.get():
                if event.type == QUIT: 
                    pygame.quit()
                    sys.exit()

    def inicio(self):
        """Muestra la imagen de nuestro videojuego con música de fondo"""  
        pygame.mixer.music.load("Sonidos/intro.mp3")
        pygame.mixer.music.play(1)
        pygame.mixer.music.set_volume(0.1)

        while True:
            screen.fill((255,255,255))
            screen.blit(self.portada,((screen.get_width()/2)-400,(screen.get_height()/2)-294))
            mensaje = fuente.render('Presione ENTER para continuar',1,(0,0,255),True)
            screen.blit(mensaje,(400,550))
            pygame.display.update()
            
            for event in pygame.event.get():
                if event.type == QUIT: 
                    pygame.quit()
                    sys.exit()

                elif event.type == KEYDOWN:

                    if event.key == K_RETURN:
                        user.introduccion()


class Usuario(object):
        
    def introduccion(self):
        """Solicita la introducción del nombre del alumno. Muestra la opción de crear usuario"""  
        clock = pygame.time.Clock()
        self.indicador = 0
        mensaje = fuente.render('Introduce el nombre del alumno:',1,(0,0,255),True) 
        mensaje2 = fuente3.render('Crear nuevo usuario',1,(0,0,0),True)
        input_box = pygame.Rect((screen.get_width()/2)-100,(screen.get_height()/2)-64,160,32)
        new_user_box = pygame.Rect((screen.get_width()/2)-115,(screen.get_height()/2)+32,230,32)
        
        color_inactive = pygame.Color('lightskyblue3')
        color_active = pygame.Color('dodgerblue2')
        color = color_inactive
        active = False
        self.text = ''
        self.done = False

        while not self.done:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    sys.exit()
                if event.type == pygame.MOUSEBUTTONDOWN:
                    # If the user clicked on the input_box rect.
                    if input_box.collidepoint(event.pos):
                        # Toggle the active variable.
                        active = not active

                    elif new_user_box.collidepoint(event.pos):
                        self.done = True
                        nuevo.crear()
                    else:
                        active = False
                    # Change the current color of the input box.
                    color = color_active if active else color_inactive
                if event.type == pygame.KEYDOWN:
                    if active:
                        if event.key == pygame.K_RETURN:
                            user.existencia()
                            self.text = ''
                        elif event.key == pygame.K_BACKSPACE:
                            self.text = self.text[:-1]
                        else:
                            self.text += event.unicode

            screen.fill((255, 255, 255))
            screen.blit(mensaje,(400,240))
            txt_surface = fuente3.render(self.text, True, color)
            # Resize the box if the text is too long.
            self.width = max(200, txt_surface.get_width()+10)
            input_box.w = self.width
            screen.blit(txt_surface, (input_box.x+5, input_box.y+5))
            pygame.draw.rect(screen, color, input_box, 2)
            pygame.draw.rect(screen,(0,255,255), new_user_box)
            screen.blit(mensaje2,(new_user_box.x+5,new_user_box.y+5))

            pygame.display.update()
            clock.tick(30)

    def existencia(self):
        """Revisa la existencia del usuario introducido en el documento Nombres de Excel """  
        nombres = pandas.read_excel('C:/Videojuego/Proyecto/Usuarios/Nombres.xlsx',header = None) 
        nombres = list(nombres.iloc[:,0])
            
        if self.text in nombres:
            self.done = True
            campo.seleccion()
        else:
            mensaje = fuente3.render('Usuario no existente',1,(255,0,0),True) 
            screen.blit(mensaje,(572,(screen.get_height()/2)-20))
            input_box = pygame.Rect((screen.get_width()/2)-100,(screen.get_height()/2)-64,self.width,32)
            pygame.draw.rect(screen,(255,0,0), input_box, 2)
            pygame.display.update()
            pygame.time.delay(1500)

    def nombre(self):
        """Recibe y regresa el nombre del alumno introducido""" 
        nombre = self.text
        self.indicador +=1
        return nombre


class CrearUsuario(object):
    
    def crear(self):
        """Crea un nuevo usuario en caso de que no haya sido de alta anteriormente """ 
        self.indicador2 = 0
        screen.fill((255,255,255))
        clock = pygame.time.Clock()
        
        mensaje = fuente.render('Introduce el nombre del nuevo alumno:',1,(0,0,255),True) 
        mensaje2 = fuente3.render('Crear',1,(0,0,0),True)
        mensaje3 = fuente3.render('Sexo',1,(0,0,0),True)
        mensaje4 = fuente3.render('M',1,(0,0,0),True)
        mensaje5 = fuente3.render('F',1,(0,0,0),True)
        mensaje6 = fuente3.render('Usuario NO valido',1,(255,0,0),True)
        mensaje7 = fuente3.render('Presione ESC para salir',1,(0,0,255),True)
        input_box = pygame.Rect((screen.get_width()/2)-100,(screen.get_height()/2)-64,160,32)
        hombre_box = pygame.Rect((screen.get_width()/2)-30,(screen.get_height()/2)+25,25,32)
        mujer_box = pygame.Rect((screen.get_width()/2)+10,(screen.get_height()/2)+25,25,32)
        create_box = pygame.Rect((screen.get_width()/2)-32,(screen.get_height()/2)+85,70,32)
        
        color_inactive = pygame.Color('lightskyblue3')
        color_active = pygame.Color('dodgerblue2')
        
        color2_inactive = pygame.Color('dodgerblue2')
        color2_active = pygame.Color('indianred3')
        
        color = color_inactive
        color2 = color2_inactive
        color3 = color2_inactive
        
        active = False
        active2 = False
        active3 = False
        self.text2 = ''
        self.sexo = ''
        self.done = False

        while not self.done:
            for event in pygame.event.get():
                if event.type == pygame.QUIT:
                    pygame.quit()
                    sys.exit()
                if event.type == pygame.MOUSEBUTTONDOWN:
                    # If the user clicked on the input_box rect.
                    if input_box.collidepoint(event.pos):
                        # Toggle the active variable.
                        active = not active

                    elif create_box.collidepoint(event.pos):

                        if self.text2 != '':
                            self.done = True
                            nuevo.excel()
                            nuevo.bienvenida()
                            campo.seleccion()
                        else:
                            screen.blit(mensaje6,((screen.get_width()/2)-(mensaje6.get_width()/2),(screen.get_height()/2)-30))
                            pygame.display.update()
                            pygame.time.delay(1500)

                    elif hombre_box.collidepoint(event.pos):
                        active2 = not active2

                    elif mujer_box.collidepoint(event.pos):
                        active3 = not active3
                        
                    else:
                        active = False
                        active2 = False
                        active3 = False
                    # Change the current color of the input box.
                    color = color_active if active else color_inactive
                    if active3 == False:
                        color2 = color2_active if active2 and not active3 else color2_inactive
                        self.sexo = 'M'
                    if active2 == False:
                        color3 = color2_active if active3 and not active2 else color2_inactive
                        self.sexo = 'F'
                if event.type == pygame.KEYDOWN:

                    if active:                             
                        if event.key == pygame.K_BACKSPACE:
                            self.text2 = self.text2[:-1]      
                        else:
                            self.text2 += event.unicode
                            
                    if event.key == K_RETURN:
                        if self.text2 != '':
                            self.done = True
                            nuevo.excel()
                            nuevo.bienvenida()
                            campo.seleccion()
                        else:
                            screen.blit(mensaje6,((screen.get_width()/2)-(mensaje6.get_width()/2),(screen.get_height()/2)-30))
                            pygame.display.update()
                            pygame.time.delay(1500)
                            
                    elif event.key == pygame.K_ESCAPE:
                        self.done = True
                        user.introduccion() 

            screen.fill((255, 255, 255))
            screen.blit(mensaje,(330,240))
            screen.blit(mensaje7,((screen.get_width())-(mensaje7.get_width())-10,10))
            txt_surface = fuente3.render(self.text2, True, color)
            # Resize the box if the text is too long.
            width = max(200, txt_surface.get_width()+10)
            input_box.w = width
            screen.blit(txt_surface, (input_box.x+5, input_box.y+5))
            pygame.draw.rect(screen, color, input_box, 2)
            pygame.draw.rect(screen,(0,255,255), create_box)
            pygame.draw.rect(screen,color2, hombre_box)
            pygame.draw.rect(screen,color3, mujer_box)
            screen.blit(mensaje2,(create_box.x+5,create_box.y+5))
            screen.blit(mensaje3,((screen.get_width()/2)-(mensaje3.get_width()/2-3),(screen.get_height()/2)-5))
            screen.blit(mensaje4,(hombre_box.x+5,hombre_box.y+5))
            screen.blit(mensaje5,(mujer_box.x+5,mujer_box.y+5))

            pygame.display.update()
            clock.tick(30)

    def excel(self):
        """Crea un libro de Excel con el nombre del usuario""" 
        nombres =pandas.read_excel('C:/Videojuego/Proyecto/Usuarios/Nombres.xlsx',header = None)
        nombres = list(nombres.iloc[:,0])
        nombres.append(self.text2)
        clave=nombres.pop(0)
        dic={clave:nombres}

        df = pandas.DataFrame(dic)
        df =df[['Nombres']]
        writer =ExcelWriter('C:/Videojuego/Proyecto/Usuarios/Nombres.xlsx') 
        df.to_excel(writer,sheet_name='Alumnos',index=False)
        writer.save()

        df2 = pandas.DataFrame({'Juego':["----------"],'Estatus':["----------"],'Tiempo':["----------"],'Real':["----------"],'Ideal':["----------"],'Fecha':["----------"]})
        df2=df2[['Juego','Estatus','Tiempo','Real','Ideal','Fecha']]
        writer = ExcelWriter('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+self.text2+'.xlsx')
        df2.to_excel(writer,sheet_name=self.text2,index=False)
        writer.save()

    def bienvenida(self):
        """Despliega un mensaje de bienvenida una vez que un nuevo usario ha sido dado de alta """ 
        screen.fill((255,255,255))
        pygame.mixer.music.stop()
        aplausos.play()
        if self.sexo == 'M':
            letra = 'o '
        elif self.sexo == 'F':
            letra = 'a '
        bienvenida = fuente4.render('Bienvenid'+letra+ self.text2,1,(255,133,15),True)
        screen.blit(bienvenida,((screen.get_width()/2)-(bienvenida.get_width()/2),(screen.get_height()/2)-(bienvenida.get_height()/2)-30))
        pygame.display.update()
        pygame.time.delay(2000)
    

    def nombre2(self):
        """Recibe y regresa el nombre del usuario que fue creado""" 
        nombre2 = self.text2
        self.indicador2+=1
        return nombre2

class Estadisticas(object):
    def escribir(self):
        """Escribe en el excel los datos obtenidos por cada uno de lo campos semánticos """ 

        resp =pandas.read_excel('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+campo.usuario()+'.xlsx',header = None)

        juego = list(resp.iloc[:,0])
        estatus = list(resp.iloc[:,1])
        tiempo = list(resp.iloc[:,2])
        real = list(resp.iloc[:,3])
        ideal = list(resp.iloc[:,4])
        fecha = list(resp.iloc[:,5])  
       
        ljuego = juego + juego.respuestas()
        lestatus = estatus + color.imagenes() + emocion.imagenes() + numero.imagenes() + animal.imagenes()

        ltiempo = tiempo + color.tiempos() + emocion.tiempos() + numero.tiempos() + animal.tiempos()
        
        lreal = real + real.respuestas2()        
        lideal = ideal + ideal.respuestas3() 

        total = color.imagenes() + emocion.imagenes() + numero.imagenes() + animal.imagenes()
        
        nfecha = []
        for i in total:
            nfecha.append("-")

        lfecha = fecha + nfecha
            
        lcolor.pop(0)
        lemocion.pop(0)
        lnumero.pop(0)
        lanimal.pop(0)
        limagen.pop(0)
        ltiempo.pop(0)
        lsesion.pop(0)
        
        dic={'Juego':lcolor,'Estatus':lemocion,'Tiempo':ltiempo,'Real':lanimal,'Ideal':limagen,'Fecha':lsesion}
        df3 = pandas.DataFrame(dic)
        df3=df3[['Juego','Estatus','Tiempo','Real','Ideal','Fecha']]
        writer =ExcelWriter('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+campo.usuario()+'.xlsx')
        df3.to_excel(writer,sheet_name=campo.usuario(),index=False)
        writer.save()
        

class CamposSemanticos(Usuario,CrearUsuario):
    def _init_(self):
        self.colores = pygame.image.load("Imagenes/Campos/colores.png")
        self.emociones = pygame.image.load("Imagenes/Campos/emociones.png")
        self.numeros = pygame.image.load("Imagenes/Campos/numeros.png")
        self.animales = pygame.image.load("Imagenes/Campos/animales.png")
        self.derecha = pygame.image.load("Imagenes/Derecha.png")
        self.izquierda = pygame.image.load("Imagenes/Izquierda.png")


    def seleccion(self):
        """Despliega la pantalla uno en la que se hace la elección del campo semántico"""
        
        pygame.mixer.music.pause()
        pygame.mixer.stop()

        try:
            self.nombre = nuevo.nombre2()
        except:
            self.nombre = user.nombre()
        try:
            if user.nombre()=='':
                self.nombre=nuevo.nombre2()
            else:
                self.nombre = user.nombre()            
        except:
            pass
        
        screen.fill((255,255,255))
        campo.mensajes()
        mensaje = fuente.render('Selecciona el campo semántico',1,(0,0,255),True)
        numeros = fuente2.render('1                                          2                                           3                                         4 ',1,(0,0,255),True)
     


        screen.blit(mensaje,(400,40))
        screen.blit(numeros,(150,550))
        screen.blit(self.colores,(30,200))
        screen.blit(self.emociones,(340,220))
        screen.blit(self.numeros,(650,200))
        screen.blit(self.animales,(960,220))
        pygame.display.update()

        izquierda_box = pygame.Rect(0,screen.get_height()/2-73,65,66)
        screen.blit(self.izquierda,(10,screen.get_height()/2-73))
     
        

        bandera = 1
        
        while bandera == 1:
            
            for event in pygame.event.get():
                if event.type == QUIT: 
                    pygame.quit()
                    sys.exit()

             

                elif event.type == KEYDOWN:
            
                    if event.key == K_1:
                        bandera = 0
                        color.eleccion()

                    elif event.key == K_2:
                        bandera = 0
                        emocion.eleccion()

                    elif event.key == K_3:
                        bandera = 0
                        numero.eleccion()

                    elif event.key == K_4:
                        bandera = 0
                        animal.eleccion()

                    elif event.key == K_s:
                        bandera = 0

                        resp =pandas.read_excel('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+campo.usuario()+'.xlsx',header = None)

                        juego = list(resp.iloc[:,0])
                        estatus = list(resp.iloc[:,1])
                        tiempo = list(resp.iloc[:,2])
                        real = list(resp.iloc[:,3])
                        ideal = list(resp.iloc[:,4])
                        fecha = list(resp.iloc[:,5])

                        sesion = time.localtime()
                        sesi = (str(sesion.tm_mday)+'/'+str(sesion.tm_mon)+'/'+str(sesion.tm_year))

                        ljuego = juego + ["----------"]
                        lestatus = estatus + ["----------"]
                        ltiempo = tiempo + ["----------"]
                        lreal = real + ["----------"]
                        lideal = ideal + ["----------"]
                        lfecha = fecha + [sesi]
                   
                                                  
                        ljuego.pop(0)
                        lestatus.pop(0)
                        ltiempo.pop(0)
                        lreal.pop(0)
                        lideal.pop(0)
                        lfecha.pop(0)
                        
                        dic={'Juego':lcolor,'Estatus':lemocion,'Tiempo':ltiempo,'Real':lanimal,'Ideal':limagen,'Fecha':lsesion}
                        df3 = pandas.DataFrame(dic)
                        df3=df3[['Juego','Estatus','Tiempo','Real','Ideal','Fecha']]
                        writer =ExcelWriter('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+campo.usuario()+'.xlsx')
                        df3.to_excel(writer,sheet_name=campo.usuario(),index=False)
                        writer.save()
                        
                        portada.inicio()
                        

    def segunda(self):
        """Despliega la pantalla dos en la que se hace la elección del campo semántico"""

        pygame.mixer.music.pause()
        pygame.mixer.stop()

        try:
            self.nombre = nuevo.nombre2()
        except:
            self.nombre = user.nombre()
        try:
            if user.nombre()=='':
                self.nombre=nuevo.nombre2()
            else:
                self.nombre = user.nombre()            
        except:
            pass
        
        screen.fill((255,255,255))
        campo.mensajes()
        mensaje = fuente.render('Selecciona el campo semántico',1,(0,0,255),True)
        numeros = fuente2.render('1',1,(0,0,255),True)
        izquierda_box = pygame.Rect(0,screen.get_height()/2-73,65,66)
        screen.blit(self.izquierda,(10,screen.get_height()/2-73))
        mensaje2=fuente3.render('Estadísticas',1,(255,255,255),True)
        estadisticas=pygame.Rect(screen.get_width()-mensaje2.get_width()-20,10,(mensaje2.get_width())+10,(mensaje2.get_height()+10))
        pygame.draw.rect(screen,(28,173,52), estadisticas)
        screen.blit(mensaje2,(estadisticas.x+5,estadisticas.y+5))
        screen.blit(mensaje,(400,40))
        screen.blit(numeros,(230,550))
        screen.blit(self.animales,(130,170))

        pygame.display.update()
        
        bandera = 1
        
        while bandera == 1:
            
            for event in pygame.event.get():
                if event.type == QUIT: 
                    pygame.quit()
                    sys.exit()

                elif event.type == pygame.MOUSEBUTTONDOWN:
                    if estadisticas.collidepoint(event.pos):
                        bandera = 0
                        estadistica.graficas()
                        estadistica.esta()

                    elif izquierda_box.collidepoint(event.pos):
                        bandera = 0
                        campo.seleccion()

                elif event.type == KEYDOWN:
            
                    if event.key == K_1:
                        bandera = 0
                        animal.eleccion()

                    elif event.key == K_s:
                        bandera = 0

                        resp =pandas.read_excel('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+campo.usuario()+'.xlsx',header = None)

                        juego = list(resp.iloc[:,0])
                        estatus = list(resp.iloc[:,1])
                        tiempo = list(resp.iloc[:,2])
                        real = list(resp.iloc[:,3])
                        ideal = list(resp.iloc[:,4])
                        fecha = list(resp.iloc[:,5])


                        sesion = time.localtime()
                        sesi = (str(sesion.tm_mday)+'/'+str(sesion.tm_mon)+'/'+str(sesion.tm_year))

                        ljuego = juego + ["----------"]
                        lestatus = estatus + ["----------"]
                        ltiempo = tiempo + ["----------"]
                        lreal = real + ["----------"]
                        lideal = ideal + ["----------"]
                        lfecha = sesiones + [sesi]
                                                  
                        ljuego.pop(0)
                        lestatus.pop(0)
                        ltiempo.pop(0)
                        lreal.pop(0)
                        lideal.pop(0)
                        lfecha.pop(0)

                        dic={'Juego':lcolor,'Estatus':lemocion,'Tiempo [s]':ltiempo,'Real':lanimal,'Ideal':limagen,'Fecha':lsesion}
                        df3 = pandas.DataFrame(dic)
                        df3=df3[['Juego','Estatus','Tiempo [s]','Real','Ideal','Fecha']]
                        writer =ExcelWriter('C:/Videojuego/Proyecto/Usuarios/Alumnos/'+campo.usuario()+'.xlsx')
                        df3.to_excel(writer,sheet_name=campo.usuario(),index=False)
                        writer.save()
                        
                        portada.inicio()

        
    def mensajes(self):
        """Genera dos mensaje que aparecen siempre en el juego una vez que se ha iniciado sesión en algún usuario. Estos son nombre del alumno y presione S para salir"""
        self.nombre.title()
        mensaje2 = fuente3.render('Usuario: ' + self.nombre,1,(0,0,255),True)
        screen.blit(mensaje2,(10,10))
        mensaje3 = fuente3.render('Presiona S para salir',1,(0,0,255),True)
        screen.blit(mensaje3,(screen.get_width()-mensaje3.get_width()-10,screen.get_height()-mensaje3.get_height()-70))
        pygame.display.update()
        

    def usuario(self):
        """Regresa el nombre del usuario"""
        usuar = self.nombre
        return usuar

"""Lista de objetos creados para cada una de las clases definidas anteriormente"""          
color = Colores()
emocion = Emociones()
animal = Animales()
numero = Numeros()
campo = CamposSemanticos()
bot = Botonera()
portada = Portada()
user = Usuario()
nuevo = CrearUsuario()
estadistica = Estadisticas()

juego = Juegos()
estatus = Estatus()
tiempo = Tiempo()
real = Real()
ideal = Ideal()
fecha = Fecha()
                                

def main():
    """Programa principal del programa"""
    screen.fill((255,255,255))
    pygame.time.delay(2000)
    portada.cargando()
    portada.inicio()
    
if _name_ == '_main_': main()