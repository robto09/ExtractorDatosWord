from docx import Document


class Extractor(object):

    def Name(self):
        listing = str(docText.split())
        a = docText.index("Yo")
        b = listing.index("portado")
        c = docText[a:b]
        c = c.replace("Yo", "")
        c = c.replace(" portador (a) de la c�dula de identidad No.01-0924-0642 en adelante y para efectos de esta carta de compromiso se denomin", "")
        return c

    def no(self):
        listing = str(docText.split())
        a = docText.index("identidad")
        b = listing.index("portador")
        c = docText[a:b]
        c = c.replace("identidad", "")
        c = c.replace(" en adelante y para efectos de esta carta de compromiso se denomin", "")
        print("")
        return c


    def Nombre(self):
        listing = str(docText.split())
        a = docText.index("Curso")
        b = docText.index("Costo")
        c = docText[a:b]
        c = c.replace("Curso:", "")
        c = c.replace("\t", "")
        print ("")
        return c

    def Costo(self):
        listing = str(docText.split())
        a = docText.index("Costo")
        b = docText.index("Tiempo")
        c = docText[a:b]
        c = c.replace("Costo:", "")
        c = c.replace("\t", "")

        return c


    def debe(self):
        listing = str(docText.split())
        a = docText.index("laborar")
        b = docText.index("Fecha")
        c = docText[a:b]
        c = c.replace("laborar con Banco Popular:", "")
        c = c.replace("\t", "")

        print("")
        return c


    def fecha(self):
        #Fetcha
        listing = str(docText.split())
        a = docText.index("curso")
        b = docText.index("Instituci�n")
        c = docText[a:b]

        c = c.replace("Fecha del curso:", "")
        c = c.replace("curso", "")


        c = c.replace("curso', 'Fecha', 'del', 'curso:', ", "")
        c = c.replace("\t", "")


        return c


    def Institucion(self):
        listing = str(docText.split())
        a = docText.index("Instituci�n")
        b = docText.index("Lugar")
        c = docText[a:b]
        c = c.replace("Instituci�n que lo imparte:", "")
        c = c.replace("\t", "")
        return c


    def Lugar(self):
        listing = str(docText.split())
        a = docText.index("Lugar")
        b = docText.index("Subsidio")
        c = docText[a:b]
        c = c.replace("Lugar:", "")
        c = c.replace("\t", "")
        return c


    def Subsidio(self):
        listing = str(docText.split())
        a = docText.index("Subsidio")
        b = docText.index("mil")
        c = docText[a:b]
        c = c.replace("Subsidio Total: ", "")
        c = c.replace("(", "")
        c = c.replace("\t", "")
        return c


    def extractor(self, FileLocation):
        if (FileLocation[-1] == "x"):
            global document
            global docText
            document = Document(FileLocation)
            docText = '\n\n'.join([
                paragraph.text for paragraph in document.paragraphs
            ])
        else:
            print ("n")

