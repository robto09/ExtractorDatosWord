from flask import Flask
from flask import render_template
from flask import request, redirect, url_for, session, flash, send_file

from extract import Extractor

from models import *
from pony.orm import *
from werkzeug import secure_filename
from collections import OrderedDict
from datetime import date
from dateutil.relativedelta import relativedelta

import hashlib
import os
import pandas as pd

app = Flask(__name__)
app.secret_key = 'A0Zr98j/3yXxjg$R~XHH!jmN]LWX/,?RT'

months = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
		  "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
		  "septiembre": 9, "octubre": 10, "noviembre": 11,
		  "diciembre": 12}

@app.route('/', methods=["GET", "POST"])
def HomePage():
	if (request.method == 'POST'):
		f = request.files['file']
		extension = f.filename.split('.')[-1]

		if extension != "docx":
			flash("Incorrecta extencion de archivo!")
			return redirect(url_for("HomePage"))
		else:
			file_name = f.filename
			file_name = file_name.replace(' ', '_')
			f.save(secure_filename(file_name))

			e = Extractor()
			e.extractor(file_name)
			NAME = e.Name().strip()[:-1]
			NO = e.no()
			NOMBRE = e.Nombre()
			COSTO = e.Costo().replace("US$", "").replace(",00", "")
			DEBE = e.debe().strip()
			FETCHA = e.fecha()
			INSTITUCION = e.Institucion()
			LUGAR = e.Lugar()
			SUBSIDIO = e.Subsidio().replace("$", "").replace(",00", "")

			year = FETCHA.strip()[-4:]
			Month = months[FETCHA.split("de ")[1].strip()]

			final = date(int(year),Month,1)+relativedelta(months=+int(DEBE[:2]))
			final = str(final.month) + "/" + str(final.year)

			all_in_one = NAME + NO + NOMBRE + COSTO + DEBE + FETCHA + INSTITUCION + LUGAR + SUBSIDIO
			all_in_one = all_in_one.encode()
			MD5 = hashlib.md5(all_in_one).hexdigest()

			mydir = '/home/docxparser'
			filelist = [ f for f in os.listdir(mydir) if f.endswith(".docx")]
			for f in filelist:
			    os.remove(os.path.join(mydir, f))

			try:
				with db_session:
					Document(MD5=MD5,
							 No_id=NO.replace("No.", ""),
							 NAME=NAME,
							 Nombre_Curso=NOMBRE,
							 Costo=COSTO,
							 Compromiso=DEBE[:2],
							 Year=year,
							 Month=str(Month),
							 Institucion=INSTITUCION,
							 Lugar=LUGAR,
							 Subsidio=SUBSIDIO,
							 Final_Compromiso=final)

					return render_template("index.html",
											NAME=NAME,
											NO=NO,
											NOMBRE=NOMBRE,
											COSTO=COSTO,
											DEBE=DEBE,
											FETCHA=FETCHA,
											INSTITUCION=INSTITUCION,
											LUGAR=LUGAR,
											SUBSIDIO=SUBSIDIO,
											MESSAGE="Fields stored to DB!")

			except Exception as e:
				return render_template("index.html",
										NAME="",
										NO="",
										NOMBRE="",
										COSTO="",
										DEBE="",
										FETCHA="",
										INSTITUCION="",
										LUGAR="",
										SUBSIDIO="",
										MESSAGE="Este archivo ya se ha almacenado en la BD, seleccione otro.")


	return render_template("index.html")

@app.route("/export")
def export():
	with db_session:
		records = Document.select().order_by(Document.No_id)[:]
		records = [OrderedDict({"ID": r.No_id, "Nombre": r.NAME, "Nombre Del Curso": r.Nombre_Curso, "Costo": r.Costo, "Compromiso": r.Compromiso, "Fecha Del Curso": r.Month + "/" + r.Year, "Instituci�n Que Imparte Curso": r.Institucion, "Lugar": r.Lugar, "Subsidio": r.Subsidio, "Compromiso Final": r.Final_Compromiso}) for r in records]

		df = pd.DataFrame.from_dict(records)
		writer = pd.ExcelWriter('/home/docxparser/mysite/exported.xlsx', engine='xlsxwriter')
		df.to_excel(writer, index=False, sheet_name='Sheet1')

		worksheet = writer.sheets['Sheet1']
		worksheet.set_column('A:A', 15)
		worksheet.set_column('B:B', 21)
		worksheet.set_column('C:C', 45)
		worksheet.set_column('D:D', 7)
		worksheet.set_column('E:E', 11)
		worksheet.set_column('F:F', 14)
		worksheet.set_column('G:G', 32)
		worksheet.set_column('H:H', 23)
		worksheet.set_column('I:I', 7)
		worksheet.set_column('J:J', 15)

		writer.save()

	return send_file("exported.xlsx", as_attachment=True)