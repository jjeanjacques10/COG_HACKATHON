import cv2
import logging as log
import datetime as dt
from time import sleep
from openpyxl import Workbook
from microsoft.reconhecimento import ReconhecimentoMicrosoft
import os.path
from PIL import Image

book = Workbook()

cascPathFace = "haarcascade_frontalface_default.xml"
cascPathFace2 = "haarcascade_frontalface_alt_tree.xml"

faceCascade = cv2.CascadeClassifier(cascPathFace)
face2Cascade = cv2.CascadeClassifier(cascPathFace2)
log.basicConfig(filename='webcam.log',level=log.INFO)

video_capture = cv2.VideoCapture(cv2.CAP_DSHOW+1)
anterior = anterior2 = 0
sheet = book.active
sheet['A1'] = "Faces"
sheet['B1'] = "Dia"
sheet['C1'] = "Hora"
sheet['D1'] = "Idade"
sheet['E1'] = "Sexo"
sheet['F1'] = "Emoção"
sheet['G1'] = "Câmera"
sheet['H1'] = "Local"
i = 2

while True:
    print("Esse é o valor do i = {}".format(i))
    celA = 'A{}'.format(i)
    celB = 'B{}'.format(i)
    celC = 'C{}'.format(i)
    celD = 'D{}'.format(i)
    celE = 'E{}'.format(i)
    celF = 'F{}'.format(i)
    celG = 'G{}'.format(i)
    celH = 'H{}'.format(i)

    if not video_capture.isOpened():
        print('Unable to load camera.')
        sleep(5)
        pass

    # Capture frame-by-frame
    ret, frame = video_capture.read()

    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

    # Detectando face
    faces = faceCascade.detectMultiScale(
        gray,
        scaleFactor=1.1,
        minNeighbors=5,
        minSize=(30, 30)
    )

    faces2 = face2Cascade.detectMultiScale(
        gray,
        scaleFactor=1.1,
        minNeighbors=5,
        minSize=(30, 30)
    )

    # Draw a rectangle around the faces
    for (x, y, w, h) in faces:
        cv2.rectangle(frame, (x, y), (x+w, y+h), (0, 255, 0), 2)
    for (x2, y2, w2, h2) in faces2:
        cv2.rectangle(frame, (x2, y2), (x2+w2, y2+h2), (0, 255, 0), 2)



    if(anterior != len(faces)):
        print(len(faces))
        print("Entrou")
        anterior = len(faces)

        path = "../faces/"
        imagem_path = path + "face{}.jpeg".format(i)
        im = Image.fromarray(frame)
        im.save(imagem_path)

        try:
            if(os.path.exists(imagem_path)):
                #Reconhecendo caracteristicas da face
                microsoft = ReconhecimentoMicrosoft()
                result = microsoft.indentificarFaceImagem(imagem_path)
                print(result)
                # Cadastrando no excel
                sheet[celA] = str(len(faces))
                sheet[celB] = "{}-{}-{}".format(dt.datetime.now().day, dt.datetime.now().month,
                                                      dt.datetime.now().year)
                sheet[celC] = "{}:{}:00".format(dt.datetime.now().hour, dt.datetime.now().minute)
                sheet[celD] = result['idade']
                sheet[celE] = result['sexo']
                sheet[celF] = result['emocao']
                sheet[celG] = 'Camera 1'
                sheet[celH] = 'Shopping Center Norte - Tv. Casalbuono, 120 - Vila Guilherme, São Paulo - SP, 02089-900'
                i = i + 1
        except Exception:
            print("Não encontrei ninguém: {}".format(Exception))
        print("Saiu")

    if(anterior2 != len(faces2)):
        print(len(faces2))
        print("Entrou")
        anterior2 = len(faces2)

        path = "../faces/"
        imagem_path = path + "face{}.jpeg".format(i)
        im = Image.fromarray(frame)
        im.save(imagem_path)

        try:
            if (os.path.exists(imagem_path)):
                # Reconhecendo caracteristicas da face
                microsoft = ReconhecimentoMicrosoft()
                result = microsoft.indentificarFaceImagem(imagem_path)
                print(result)
                # Cadastrando no excel
                sheet[celA] = str(len(faces))
                sheet[celB] = "{}-{}-{}".format(dt.datetime.now().day, dt.datetime.now().month,
                                                      dt.datetime.now().year)
                sheet[celC] = "{}:{}:00".format(dt.datetime.now().hour, dt.datetime.now().minute)
                sheet[celD] = result['idade']
                sheet[celE] = result['sexo']
                sheet[celF] = result['emocao']
                sheet[celG] = 'Camera 1'
                sheet[celH] = 'Shopping Center Norte - Tv. Casalbuono, 120 - Vila Guilherme, São Paulo - SP, 02089-900'
                i = i + 1
        except Exception:
            print("Não encontrei ninguém: {}".format(Exception))
        print("Saiu")

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

    # Salvando arquivo excel
    book.save('../resultados.xlsx')

# When everything is done, release the capture
video_capture.release()
cv2.destroyAllWindows()
