from art import *
import os

abecedario = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "Ñ", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    RED = '\033[31m'

def interface():
    print(bcolors.BOLD + bcolors.OKGREEN)
    tprint("toAiken")
    tprint("by Andressance", font="small")
    print(bcolors.BOLD + bcolors.RED + "Bienvenido al programa de creación de documentos txt en formato Aiken \n")
    print("Funcionamiento: \n" + bcolors.ENDC+ bcolors.OKGREEN)
    print("El programa transformará el documento word a txt, añadiendo la respuesta correcta a cada pregunta si está en negrita o subrayada en el documento word")
    print("Introduce el nombre del documento")
    print("Introduce el número de preguntas que quieres crear, en caso de que quieras hacer menos de las que has escrito escribe 'SALIR' y el documento se creará con las preguntas que hayas introducido")
    print("Tras introducir la pregunta, escribe el número de respuestas que quieres que tiene y posteriormente su contenido, finalmente se te pedirá la respuesta correcta")

def createQuestions():

    text = []
    numQuest = int(input("Introduce el número de preguntas a crear: "))
    for i in range(numQuest):
        pregunta = str(input(f"Introduce la pregunta {i + 1}: "))
        text.append(pregunta)
        text.append("\n")
        numAnsw = int(input("Introduce el número de respuestas de la pregunta: "))
        for j in range(numAnsw):
            res = abecedario[j] + ". "
            answ = str(input(f"Introduce la respuesta {i + 1}: " ))
            res += answ
            text.append(res)
            text.append("\n")
        correctAnsw = "ANSWER: "
        char = str(input("Introduce la respuesta correcta: "))
        char = char.strip().upper()
        correctAnsw += char
        text.append(correctAnsw)
        text.append("\n")

    return text

def createTxt(name:str, text:list):
    
    if ".txt" in name:
        with open(docname, "w") as f:
            for line in text:
                f.write(line)
            f.close()

        os.startfile(docname)
    else:   
        docname = name + ".txt"
        with open(docname, "w") as f:
            for line in text:
                f.write(line)
            f.close()

    os.startfile(docname)

def main():
    interface()
    name = str(input("\n Introduce el nombre del archivo a crear: "))
    createTxt(name ,createQuestions())

main()