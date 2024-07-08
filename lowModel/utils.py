import os
from termcolor import colored
from msvcrt import kbhit, getch
from unidecode import unidecode

def center(text:str,fill=' '):#Centraliza um texto no terminal
    terminalWidth = os.get_terminal_size()[0]#Tamanho do terminal

    if '\033' in text:#Caso o texto esteja colorido
        colors = []
        uncoloredTexts = []
        while '\033' in text:#Enquanto existirem partes coloridas
            #Separa a cor
            colorBegin = text.find('\033')
            colorEnd = text.find('m',colorBegin) + 1
            while text[colorEnd] == '\033':#Enquanto uma cor terminar com o início de uma nova cor, adiciona
                colorEnd = text.find('m',colorEnd) + 1
            colors.append(text[colorBegin:colorEnd])

            #Separa a versão do texto não colorida
            uncoloredTexts.append(text[colorEnd:text.find('\033[0m')])

            #Remove a cor salva
            text = text.replace(colors[-1],'',1).replace('\033[0m','',1)
        
        #Quando o texto estiver sem cor
        text = text.center(terminalWidth,fill)#Centraliza o texto
        result = ''
        for n,uncoloredText in enumerate(uncoloredTexts):#Pra cada cor removida
            #Colore o texto novamente e guarda a posição onde ela está depois de ser centralizada
            coloredText = colors[n] + uncoloredText + '\033[0m'
            uncoloredBegin = text.find(uncoloredText)
            uncoloredEnd = uncoloredBegin + len(uncoloredText)

            #Vai adicionando ao resultado o mesmo texto centralizado porém substituindo a parte não colorida pela colorida
            result += text[:uncoloredBegin] + coloredText
            text = text[uncoloredEnd:]
        return result + text#Retorna o texto centralizado sem modificar as cores
    else:#Caso não esteja colorido, retorna o valor centralizado normalmente
        return text.center(terminalWidth,fill)

def simplifyText(text:str):
    return unidecode(text.lower())

def numToMonth(num,upper=True):
    months = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro']
    num = int(num) - 1
    if num > 0 and num < len(months):
        return months[num].upper() if upper else months[num]

def numToMoney(num):
    text = f'{num:.2f}'.replace('.','')
    if len(text) >= 3:
        text = text[:-2] + ',' + text[-2:]

    for i in range(100):
        if len(text) > 6+4*i:
            text = text[:-6-4*i] + '.' + text[-6-4*i:]
        else:
            break

    return 'R$ ' + text

def moneyToNum(text):
    text = text.replace('R$ ','')
    if text[-3:] == ',00':
        text = text[:-3]
    return float(text.replace(',','.').replace('.',''))

class Pointer:
    #Ponteiros são espécies de "locais na memória compartilhados"
    #No momento que um ponteiro recebe outro ponteiro como valor eles são conectados
    #Uma vez conectados os ponteiros compartilham sempre o mesmo valor, alterar um resulta na alteração do outro

    def __init__(self,value=None):#Transforma um valor em um ponteiror
        self.list = [value]

    def get(self):#Retorna o valor do ponteiro
        if type(self.list[0]) == Pointer:
            return self.list[0].get()
        return self.list[0]
    
    def set(self,value):#Edita o valor do ponteiro
        if type(self.list[0]) == Pointer:
            self.list[0].set(value)
        else:
            self.list = [value]
    
    def add(self,value):#Adiciona ao valor atual do ponteiro um novo valor (pode somar strings numéricas)
        if type(self.get()) == str and self.get().isnumeric():
            self.set(str(int(self.get())+value))
            return 0
        self.set(self.get()+value)

    def append(self,value):#Caso o ponteiro aponte para uma lista, insere um novo elemento na última posição
        if type(self.get()) == list:
            oldList = self.get()
            oldList.append(value)
            self.set(oldList)

    def insert(self,index,value):#Caso o ponteiro aponte para uma lista, insere um novo elemento na posição index
        if type(self.get()) == list:
            oldList = self.get()
            if index == -1:
                oldList.append(value)
            else:
                oldList.insert(index,value)
            self.set(oldList)

    def remove(self,value):#Caso o ponteiro aponte para uma lista, remove um elemento por seu valor
        if type(self.get()) == list:
            oldList = self.get()
            oldList.remove(value)
            self.set(oldList)

    def pop(self,index=-1):#Caso o ponteiro aponte para uma lista, remove um elemento pela sua posição
        if type(self.get()) == list:
            oldList = self.get()
            oldList.pop(index)
            self.set(oldList)

    def sort(self,reverse=False,key=None):#Caso o ponteiro aponte para uma lista, organiza
        #Uma key é uma função que recebe um valor e retorna um valor numérico, a lista será organizada com base neste valor
        if type(self.get()) == list:
            oldList = self.get()
            oldList.sort(reverse,key)
            self.set(oldList)

    def __str__(self):#Permite que ponteiros sejam tratados como string
        return str(self.get())

class Keyboard:
    #Filters
    NUMBERS = '1234567890'

    def readKeyboard(orded=True) -> list:
        keys = []
        while kbhit():
            if orded:
                keys.append(ord(getch()))
            else:
                keys.append(getch())
        return keys

    def keyToAccentedChar(keyNumber):#Recebe o código de uma tecla do teclado e retorna a letra acentuada correspondente
        base = 'Ç éâ à çê è îì  É  ô òûù        áíóú                 ÁÂÀ              ãÃ          Ê È ÍÎ      Ì Ó ÔÒõÕ   ÚÛÙ'
        try:
            return base[keyNumber-128]
        except:
            return keyNumber
        
    def keyToSpecialsChar(keyNumber):
        specials = {32 : 'space',27 : 'esc',13 : 'enter',9 : 'tab',8 : 'back'}
        try:
            return specials[keyNumber]
        except:
            return keyNumber
        
    def keyToSpecials224Char(keyNumber:int) -> str:
        specials224 = {72 : 'up',80 : 'down',75 : 'left',77 : 'right',83 : 'delete'}
        try:
            return specials224[keyNumber]
        except:
            return 'error'

    def convertKeys(keys:list) -> list:
        convertedKeys = []
        special224 = False
        for key in keys:
            if key == 224:#Verifica se é uma tecla que começa com 224 (ex: [224,71] = 'up')
                special224 = True
            elif special224:#Se for uma tecla com 224 a próxima tecla será analizada de forma diferente
                special224 = False
                convertedKeys.append(Keyboard.keyToSpecials224Char(key))
            else:#Caso não seja 224, verifica se é uma tecla normal, caso não, converte para acentuada ou especial
                if key < 128:
                    newkey = Keyboard.keyToSpecialsChar(key)
                    if key != newkey:
                        convertedKeys.append(newkey)
                    else:
                        convertedKeys.append(chr(newkey))
                else:
                    convertedKeys.append(Keyboard.keyToAccentedChar(key))
        return convertedKeys

    def getKeyPressed() -> list:#Retorna a tecla pressionada
        keys = Keyboard.convertKeys(Keyboard.readKeyboard())
        if len(keys) == 1:
            return keys[0]
        return keys
    
