import sys
import os
from flask import Flask, request, render_template, jsonify, make_response, send_file, flash, url_for, redirect
import pyexcel
import HTML
import pdfkit
import zipfile
import datetime

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST' and 'excel' in request.files:
        try:
	        filename = request.files['excel'].filename
	        extension = filename.split(".")[1]
	        content = request.files['excel'].read()
	        numbers = request.form.getlist('accounts')
	        accountlist = numbers[0].splitlines()
	        striped_accounts = [i.strip() for i in accountlist]
	        accountlist = striped_accounts
	        styling = 'display: inline; page-break-before: auto; padding-bottom: 50%; font-family: Calibri; font-size: 8.76;'
	        wb = pyexcel.get_book(file_type=extension, file_content=content)
	        sheets = wb.to_dict()
	        all_sheets = []
	        for name in sheets.keys():
	            all_sheets.append(name)
	        filestozip = []
	        for k in all_sheets:
	            for i in range(1,len(wb[k].column[1])):
	                if len(accountlist) == 0:
	                    return render_template('upload.html')
	                wbname = pyexcel.Sheet([wb[k].row[i]])
	                if str(wbname[0,0]).strip() in accountlist:
	            		    htmlcode = HTML.table()
	            		    array = [wb[k].row[0],wb[k].row[i]]
	            		    filestozip.append(str(wbname[0,0]).strip())
	            		    for j in range(0,len(array[0])):
	            		        if ("ssn" in str(array[0][j]).lower() or "tax" in str(array[0][j]).lower() or "social" in str(array[0][j]).lower()) and (len(str(array[1][j])) != "0" or len(str(array[1][j])) != "1"):
	            		            if len(str(array[1][j])) == 8:
	            		                htmlcode += HTML.table([[str(array[0][j])],[str('XXX-XX-X' + str(array[1][j])[5:])]],border=0,style=(styling))
	            		            elif len(str(array[1][j])) == 0 or len(str(array[1][j])) == 1:
	            		                htmlcode += HTML.table([[str(array[0][j])],[str(array[1][j]).strip()]],border=0,style=(styling))
	            		            else:
	            		                htmlcode += HTML.table([[str(array[0][j])],[str('XXX-XX-X' + str(array[1][j])[6:])]],border=0,style=(styling))
	            		        elif len(str(array[1][j])) == 9:
	            		            htmlcode += HTML.table([[str(array[0][j])],[str(str(array[1][j])[0:5] + "-" + str(array[1][j])[5:])]],border=0,style=(styling))
	            		        elif len(str(array[1][j])) == 10:
	            		            if "ph" in str(array[0][j]).lower() and array[1][j] > 10000000:
	            		                htmlcode += HTML.table([[str(array[0][j])],[str("(" + str(array[1][j])[0:3] + ") " + str(array[1][j])[3:6] + "-" + str(array[1][j])[6:])]],border=0,style=(styling))
	            		            else:
	            		                htmlcode += HTML.table([[str(array[0][j])],[str(array[1][j]).strip()]],border=0,style=(styling))
	            		        elif "email" in str(array[0][j]).lower():
	            		        		if len(str(array[1][j])) >= 1:
	            		        			htmlcode += HTML.table([[str(array[0][j])],[str('XXXXX')]],border=0,style=(styling))
	            		        		else:
	            		        			htmlcode += HTML.table([[str(array[0][j])],[str(array[1][j]).strip()]],border=0,style=(styling))
	            		        elif "sale_price" in str(array[0][j]).lower() or "proceeds" in str(array[0][j]).lower():
	            		        		if len(str(array[1][j])) >= 1:
	            		        			htmlcode += HTML.table([[str(array[0][j])],[str('XXXXX')]],border=0,style=(styling))
	            		        		else:
	            		        			htmlcode += HTML.table([[str(array[0][j])],[str(array[1][j]).strip()]],border=0,style=(styling))
	            		        #elif "acct" in str(array[0][j]).lower() or "account" in str(array[0][j]).lower():
	            		        else:
	            		            htmlcode += HTML.table([[str(array[0][j])],[str(array[1][j]).strip()]],border=0,style=(styling))
	            		    f = open('/' + str(i) + '.html', 'w')
	            		    f.write(htmlcode)
	            		    f.close()
	            		    pdfkit.from_file('/' + str(i) + '.html', '/' + str(wbname[0,0]).strip() + '.pdf', options={'orientation': 'Landscape'})
	            		    os.remove('/' + str(i) + '.html')
	        if len(filestozip) > 0:
	            newzip = zipfile.ZipFile('/' + str(filename.split(".")[0]) + '.zip', mode='w')
	        for i in filestozip:
	            newzip.write('/' + str(i).strip() + '.pdf')
	            os.remove('/' + str(i).strip() + '.pdf')
	        if len(filestozip) > 0:
	            newzip.close()
	            return send_file(filename_or_fp='/' + str(filename.split(".")[0]) + '.zip',attachment_filename=str(filename.split(".")[0]) + '.zip', as_attachment=True)
        except:
            return render_template('upload.html')
    return render_template('upload.html')
    
if __name__ == "__main__":
    # start web server
    app.run(
        host="0.0.0.0",
        port=int("80"),
        debug=True
    )