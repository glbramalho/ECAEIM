#!/usr/bin/python
# coding: latin-1

#
# Geraldo Ramalho jun/2013
# v0.0 - geracao automatica do PDF do PPC para o estilo ppc_utf8.sty

# import time
import subprocess
# from unicodedata import normalize
# import csv
# from xlrd import open_workbook
# import unicodedata
# from datetime import date, datetime, timedelta
# import codecs

# nome do arquivo TEX principal
arquivoTEX='PPC-ECAEIM-IFCE-Maracanau'
arquivoFlag='PPC-flag.tex'

cursos=[0,1]; # 0=ECA; 1=EIM
for n in cursos:
	if n==0:
		nomecurso='ECA' 
	else:
		nomecurso='EIM'
	#
	f = open(arquivoFlag,'w')
	f.write('\\def\\COMPILAR{%s}\n'%n)
	f.close()
	#
	print nomecurso
	#
	subprocess.call(['pdflatex','-interaction=nonstopmode', arquivoTEX+'.tex'])
 	subprocess.call(['pdflatex','-interaction=nonstopmode', arquivoTEX+'.tex'])
	subprocess.call(['mv', arquivoTEX+'.pdf', 'PPC-'+nomecurso+'.pdf'])
	subprocess.call(['rm', arquivoTEX+'.log'])
	subprocess.call(['rm', arquivoTEX+'.out'])
	subprocess.call(['rm', arquivoTEX+'.aux'])
	subprocess.call(['rm', arquivoTEX+'.toc'])
	subprocess.call(['rm', arquivoTEX+'.bak'])


