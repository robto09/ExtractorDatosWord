from pony.orm import *

db = Database()

class Document(db.Entity):
	MD5=Required(str, unique=True)
	No_id=Required(str)
	NAME=Required(str)
	Nombre_Curso=Required(str)
	Costo=Required(str)
	Compromiso=Required(str)
	Year=Required(str)
	Month=Required(str)
	Institucion=Required(str)
	Lugar=Required(str)
	Subsidio=Required(str)
	Final_Compromiso=Required(str)

db.bind(provider='sqlite', filename='documents.db', create_db=True)
db.generate_mapping(create_tables=True)