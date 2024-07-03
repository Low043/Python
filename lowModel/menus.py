from lowModel.utils import *
from datetime import date, timedelta

class MenuOption:
    def __init__(self,text):
        self.menu = Pointer()
        self.text = Pointer(text)
        self.editVar = Pointer()
        self.filter = Pointer()
        self.rule = Pointer(lambda *args : None)#Regrinha do que acontece depois da interação (é útil com data)
        self.enterFunctionList = Pointer([])
        self.validInteraction = False

    def synchronizeValue(self,pointer:Pointer,filter=None,rule=lambda *args : None):
        self.editVar.set(pointer)
        self.filter.set(filter)
        self.rule.set(rule)

    def addEnterFunction(self,func):
        self.enterFunctionList.append(func)

    def closeMenu(self):#Só fecha o último Menu associado a essa opção (certifique-se que está aberto)
        self.menu.get().closeMenu.set(True)

    def interact(self,key) -> bool:
        if key == 'space':
            key = ' '
        if key == 'enter':
            for func in self.enterFunctionList.get():
                func()
                self.validInteraction = True
        elif self.editVar.get() != None:
            if key == 'back':
                self.editVar.set(self.editVar.get()[:-1])
                self.validInteraction = True
            elif len(key) == 1:
                if self.filter.get() == None or key in self.filter.get():
                    self.editVar.set(self.editVar.get()+key)
                    self.validInteraction = True
        self.rule.get()(key)
        return self.validInteract()

    def validInteract(self) -> bool:
        if self.validInteraction:
            self.validInteraction = False
            return True
        return False

    def ruleDate(self,key):
        text = self.editVar.get()
        size = len(text)
        if size > 10:
            self.editVar.set(text[:-1])
            self.validInteraction = False
        elif size == 3 or size == 6:
            if key == 'back':
                self.editVar.set(text[:-1])
            else:
                self.editVar.set(text[:size-1]+'/'+text[size-1:])
        elif size == 10:
            if key == 'right':
                add = 1
            elif key == 'left':
                add = -1
            else:
                return 0
            actualDate = self.editVar.get()
            convertedDate = date(int(actualDate[6:]),int(actualDate[3:5]),int(actualDate[0:2]))
            nextDate = convertedDate + timedelta(days=add)
            convertedNextDate = nextDate.strftime('%d/%m/%Y')
            self.editVar.set(convertedNextDate)
            self.validInteraction = True
            
    def __str__(self):
        result = self.text.get()
        if self.editVar.get() != None:
            result += f': {self.editVar.get()}'
        return result
    
class CheckBox(MenuOption):
    def __init__(self,text,pointer:Pointer):
        super().__init__(text)
        self.editVarBool = Pointer(pointer)
        self.editVarBool.set(bool(self.editVarBool.get()))
        self.addEnterFunction(self.check)
    
    def check(self):
        self.editVarBool.set(not self.editVarBool.get())
        self.validInteraction = True

    def __str__(self):
        return f'{"[OK]" if self.editVarBool.get() else "[  ]"} {self.text.get()}'

class SelectOption(MenuOption):
    def __init__(self,text,pointer:Pointer,optionList,selectedOption=0,attributeText=None):
        super().__init__(text)
        self.editPointer = Pointer(pointer)
        self.optionList = Pointer(optionList)
        self.selectedOption = Pointer(selectedOption)
        self.editPointer.set(self.optionList.get()[self.selectedOption.get()])
        self.rule = Pointer(self.move)
        self.attributeText = attributeText

    def move(self,key):
        move = (key == 'right' and self.selectedOption.get() < len(self.optionList.get())-1) - (key == 'left' and self.selectedOption.get() > 0)
        if move:
            self.selectedOption.add(move)
            self.editPointer.set(self.optionList.get()[self.selectedOption.get()])
            self.validInteraction = True

    def __str__(self):
        text = self.optionList.get()[self.selectedOption.get()]
        if self.attributeText:
            text = getattr(text,self.attributeText)
        if self.editVar.get() != None:
            return f'{self.text.get()}: {self.editVar.get()} {"<" if self.selectedOption.get() > 0 else " "} {text} {">" if self.selectedOption.get() < len(self.optionList.get())-1 else " "}'
        return f'{self.text.get()}: {"<" if self.selectedOption.get() > 0 else " "} {text} {">" if self.selectedOption.get() < len(self.optionList.get())-1 else " "}'

class Menu:
    def __init__(self,title,subtitle=None):
        self.title = Pointer(title)
        self.subtitle = Pointer(subtitle)
        self.optionSelected = Pointer(0)
        self.optionList = Pointer([])
        self.closeMenu = Pointer(False)
        self.overlay = False

    def main(self):#Menu principal que faz o menu funcionar
        self.show()
        while self.closeMenu.get() == False:
            try:
                self.navigate()
            except Exception as e:
                print(e)
        self.closeMenu.set(False)

    def navigate(self) -> bool:#Navega pelo menu (retorna 1 para fechar o menu)
        key = Keyboard.getKeyPressed()#Descobre qual tecla foi pressionada
        if key != None:#Caso tenha sido pressionada uma tecla
            if key == 'esc':#Caso seja apertado ESC, é retornado um sinal para fechar o menu
                self.closeMenu.set(True)
                return 1
            #Checa se foi pressionada alguma tecla de movimento entre opções e realiza a movimentação
            move = (key=='down') - (key=='up')
            if move:
                self.optionSelected.set((self.optionSelected.get() + move) % len(self.optionList.get()))
                self.show()
            else:#Caso nada disso tenha ocorrido, a opção selecionada decidirá o que fazer com essa tecla
                if type(key) != list:
                    key = [key]
                interactions = []
                for k in key:
                    interactions.append(self.optionList.get()[self.optionSelected.get()].interact(k))
                if True in interactions:
                    self.show()

    def addOption(self,option:MenuOption,index=-1):
        self.optionList.insert(index,option)
        option.menu.set(self)

    def removeOption(self,option:MenuOption):
        index = self.optionList.get().index(option)
        self.optionList.remove(option)
        if self.optionSelected.get() >= len(self.optionList.get()):
            self.optionSelected.add(-1)
        self.show()
        return index

    def show(self):
        titles = 1
        lines = [center(self.title.get())]

        if self.subtitle.get():
            for subtitle in self.subtitle.get().split('\n'):
                lines.append(center(subtitle))
                titles += 1
        
        for n,option in enumerate(self.optionList.get()):
            option = str(option)
            if self.overlay and n == self.optionSelected.get():
                option = colored(option,attrs=['underline'])
            lineSep = option[0] == '\n'
            lines.append(f'{"\n" if lineSep else ""}{" > " if self.optionSelected.get()==n else "   "}{option if lineSep==False else option[1:]}')
        
        maxLines = os.get_terminal_size()[1] - 2
        currentLine = titles + self.optionSelected.get()
        start = max(0,currentLine-maxLines/2)
        end = min(len(lines),maxLines+start)
        
        os.system('cls')
        for n,line in enumerate(lines):
            if n >= start and n <= end:
                print(line)

    def clearOptions(self):
        self.optionList = Pointer([])
