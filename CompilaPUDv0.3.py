#!/usr/bin/python
# coding: latin-1

#
# Geraldo Ramalho abr/2013
# v0.0 - geracao automatica dos PUDs para o estilo pud_utf8.sty a partir de arquivos CSV
# v0.1 - geracao automatica dos PUDs para o estilo pud2_utf8.sty que melhora o problema de quebra de pagina do conteudo
# v0.11 - correcao de bug que fazia desaparecer a primeira referencia da bibliografia
# v0.12 - acrescentado totais de disciplinas e totais de iCH por categoria: basica, profissionalizante e total; inclui disciplinas que são pre-requisitos
# v0.13 - ajustes para incluir automaticamente as disciplinas do campo pre-requisitos quando semestre=99
# v0.14 - ajustes de layout e geração de dois PDFs, um para cada curso
# v0.15 - gera a matriz em arquivo TEX separado
# v0.2 - Le dados diretamente do arquivo .xlsx (requer pacote xlrd) - mai/2013
# v0.21 - Inclui # para indicar livros da biblioteca virtual - mai/2013
# v0.22 - correcoes nas listas de pendencias - mai/2013
# v0.23 - substitui AC9xx por ACE0xx; opção de não gerar relatório e remover o (ND-xxx) da bibliografia; não lista Tópicos Especiais - jun/2013
# v0.24 - não listar PUDs de disciplinas com 0 créditos!!! - jun/2013
# v0.3 - correção de bug na ordenação das disciplinas nos semestres - jun/2013

import time
import subprocess
from unicodedata import normalize
import csv
from xlrd import open_workbook
import unicodedata
from datetime import date, datetime, timedelta
import codecs


# Busca do ISBN pelo título
# http://www.bn.br/site/pages/servicosProfissionais/agenciaISBN/isbnBusca/FbnBuscaISBNCatalogo.asp?pIntPagina=1&pField=Titulo&pTitulo=Metodologia do trabalho&hidbTitulo=true
# Busca do ISBN pelos 8 ultimos digitos
# http://www.bn.br/site/pages/servicosProfissionais/agenciaISBN/isbnBusca/FbnBuscaISBNCatalogo.asp?pField=ISBN&pIntPagina=1&pISBN=76050490&nidbISBN=true&hidbISBN=true

# Burca isbn em uma string
# import re
# 
# isbn = re.compile("(?:[0-9]{3}-)?[0-9]{1,5}-[0-9]{1,7}-[0-9]{1,6}-[0-9]")
# 
# matches = []
# 
# with open("text_isbn") as isbn_lines:
#     for line in isbn_lines:
#         matches.extend(isbn.findall(line))
        

def check_digit_10(isbn):
    assert len(isbn) == 9
    sum = 0
    for i in range(len(isbn)):
        c = int(isbn[i])
        w = i + 1
        sum += w * c
    r = sum % 11
    if r == 10: return 'X'
    else: return str(r)

def check_digit_13(isbn):
    assert len(isbn) == 12
    sum = 0
    for i in range(len(isbn)):
        c = int(isbn[i])
        if i % 2: w = 3
        else: w = 1
        sum += w * c
    r = 10 - (sum % 10)
    if r == 10: return '0'
    else: return str(r)

def convert_10_to_13(isbn):
    assert len(isbn) == 10
    prefix = '978' + isbn[:-1]
    check = check_digit_13(prefix)
    return prefix + check
    

def separalivros(colr2,semstr,sCodCompleto,sNom,list):
	num=colr2.count('&')
	fim=0
	while num>0:
		inix=colr2.find(' &',fim)
		ini=inix+1
		fim=colr2.find(' &',ini)
		if inix>=0:
			livro='S'+semstr+' \\hyperref[PUD:%s]{%s}'%(sCodCompleto.rstrip(),sCodCompleto.rstrip())+' '+sNom+'\n\n'+colr2[ini:fim]
			livro=livro.replace('\\\\','')
			livro=prepara_biblinkisbn(livro)
			# verifica se livro já existe na lista
			existe=0
			for x in list:
				if x.find(livro[1:150])>=0:
					existe=1
			if existe==0:
				list.append(livro)
		num-=1
		if fim<ini:
			fim=-1
	return list
    
def analisa(teste,colr2,semstr,sCodCompleto,sNom,list):
	if colr2.find(teste)>=0:
		num=colr2.count('&')
		fim=0
		while num>0:
			inix=colr2.find(teste,fim)
			ini=inix+13
			fim=colr2.find('&',ini)
			if inix>=0:
				livro='S'+semstr+' \\hyperref[PUD:%s]{%s}'%(sCodCompleto.rstrip(),sCodCompleto.rstrip())+' '+sNom+'\n\n'+colr2[ini:fim]
				livro=livro.replace('\\\\','')
				livro=prepara_biblinkisbn(livro)
				list.append(livro)
			num-=1
			if fim<ini:
				fim=-1
	return list


def ajustadata(d):
	from1900to1970 = datetime(1970,1,1) - datetime(1900,1,1) + timedelta(days=2)
	d = date.fromtimestamp( int(d) * 86400) - from1900to1970
	return d

def remove_accents(s):
	nkfd_form = unicodedata.normalize('NFKD', unicode(s))
	s = nkfd_form.encode('ASCII', 'ignore')
	return s									  

def codif(s):
	if len(s)>0:
 		s=s.encode('utf_8');
	return s
    
def corrige_pontuacao(col):
	col=col.strip()
	if len(col)>0:
		col=col.replace('.','. ')
		col=col.replace(' .','.')
		col=col.replace(';','; ')
		col=col.replace(' ;',';')
		col=col.replace(' :',':')
		col=col.replace('&','\\&')
		col=col.strip()
		if col[len(col)-1]==' ':
			col[len(col)-1]=='.'
		if col[len(col)-1]==';':
			col[len(col)-1]=='.'
		if col[len(col)-1]!='.':
			col=col+'.'
		col=col.replace(';.','.')
	return col
	
def prepara_tabela(colr):
	colr=colr.replace('*','\\\\ \n & ')+' '
	if colr.find('\\\\')==0:
		colr=colr[colr.find('\\\\')+4:]
	return colr
	
def prepara_bib(colr):
	colr2= colr.replace('@','\\\\ \n & (ND-ADQUIRIR) ')
	colr2=colr2.replace('$','\\\\ \n & (ND-COMPRADO) ')
	colr2=colr2.replace('#','\\\\ \n & (ND-BIB.VIRT) ')
	colr2=colr2.replace('ISBN  ','ISBN ')
	colr2=colr2.replace('ISBN: ','ISBN ')
	colr2=colr2.replace('ISBN:','ISBN ')
	colr2=colr2.replace('ISBN= ','ISBN ')
	colr2=colr2.replace('ISBN=','ISBN ')
	posit=colr2.find('ISBN')
# 	print '>>>>>>>>>>>>>>>%s'%colr2
	while posit>=0:
		posit=colr2.find('ISBN ',posit)
		if posit>=0:
			posit+=5
			isbnorig=colr2[posit:posit+17]
			p=0;
			while p<len(isbnorig) and '01234567890X-'.find(isbnorig[p])>=0:
				p+=1
				posit+=1
			if p>=10:	
				isbnorig=isbnorig[0:p]	
				isbn=isbnorig.strip()
				isbn=isbn.replace('.','')
				isbn=isbn.replace('-','')
				isbn=isbn.replace(' ','')
				isbn=isbn[0:14]
				isbn=isbn.strip()
				print 'isbnorig=%s'%isbnorig
				print 'isbn=%s'%isbn
				print 'tam=%d'%len(isbn)
				posit=posit+len(isbn)
				#
				if len(isbn)<13:
					print '>>>>>>>>>>isbn a converter: %s'%isbn[0:10]
					dig=check_digit_10(isbn[0:9])	
					if dig==isbn[9]:
						print 'isbn10 ok'
					else:
						print 'isbn10 %s invalido (dig não é %s)!!!!!!!!!!!!!!!!!!!'%(isbn,dig)
						isbn=isbn[0:9]+dig
					isbn=convert_10_to_13(isbn[0:10])
					print isbn
				dig=check_digit_13(isbn[0:12])	
				if dig==isbn[12]:
					print 'isbn13 ok'
				else:
					print 'isbn13 %s será corrigido para dig=%s!!!!!!!!!!!!!!!!!!!'%(isbn,dig)
					isbn=isbn[0:12]+dig
# 				isbnlnk='\\href{http://www.bn.br/site/pages/servicosProfissionais/agenciaISBN/isbnBusca/FbnBuscaISBNCatalogo.asp?pField=ISBN\\&pIntPagina=1\\&pISBN=%s\\&nidbISBN=true\\&hidbISBN=true}{%s}'%(isbn[5:13],isbn)
#  				isbnlnk='\\href{http://www.bn.br/site/pages/servicosProfissionais/agenciaISBN/isbnBusca/FbnBuscaISBNCatalogo.asp?pField=ISBN&pIntPagina=1&pISBN=%s&nidbISBN=true&hidbISBN=true}{%s}'%(isbn[5:13],isbn)
 				isbnlnk=isbn
				colr2=colr2.replace(isbnorig,isbnlnk)
# 				print colr2
# 				exit()
				# debug
# 				if colr2.find('RAMALHO')>=0:
# 					print isbnorig
# 					print isbn
# 					print colr2
# 					exit()			
	if colr2.find('\\\\')==0:
		colr2=colr2[colr2.find('\\\\')+4:]
	return colr2

def prepara_biblinkisbn(colr):
	colr2=colr
	posit=colr2.find('ISBN')
 	print '>>>>>>>>>>>>>>>%s'%colr2
	while posit>=0:
		posit=colr2.find('ISBN ',posit)
		if posit>=0:
			posit+=5
			isbn=colr2[posit:posit+13]
			isbn=isbn[0:14]
			isbn=isbn.strip()
			posit=posit+len(isbn)
			#
			if isbn.find('97885')==0: # somente brasil
				isbnlnk='\href{http://www.bn.br/site/pages/servicosProfissionais/agenciaISBN/isbnBusca/FbnBuscaISBNCatalogo.asp?pField=ISBN&pIntPagina=1&pISBN=%s&nidbISBN=true&hidbISBN=true}{%s}'%(isbn[5:13],isbn)
			elif isbn.find('9780')==0:
				isbnlnk='\href{http://www.isbnsearch.org/isbn/%s}{%s}'%(isbn,isbn)
			else:
				isbnlnk='\href{http://www.google.com/search?hl=pt-BR&tbo=p&tbm=bks&q=isbn:%s&num=10}{%s}'%(isbn,isbn)
				
			colr2=colr2.replace(isbn,isbnlnk)
	return colr2
	
	
	
# pastain="/Users/glbramalho/Downloads/"
# pastaout="../PUD'S Eng Cont Automacao/PUD automático/"
#pastain="/Users/glbramalho/Downloads/"
pastain="./"
pastaout=""#"./Latex/"
removerND=1 # 0=não remove mensagem (ND-xxx); 1= remove mensagem da bibliografia

# nome completo do cursos
CURSOS=['Engenharia Industrial Mecânica','Engenharia de Controle e Automação']


planilha=pastain+'PUDs.xlsx'
# gera um PDF para cada curso
for nomecurso in CURSOS:
	print nomecurso
	#
	codcurso = ''
	nc=nomecurso.replace(' de ',' ')
	nc=nc.replace(' e ',' ')
	codcurso=''.join([c[0].upper() for c in nc.split()])
	print codcurso
	

	header=[ 'codigo','disciplina','semestre','horas','creditos','prerequisito','objetivos','ementa','conteudo','bibbasica','bibcomplementar','autor','data','revisao','revisor','datarevisao']
	numpud=0
	listcompleta=[] # todos os livros
	list=[]
	listc=[]
	listcomprados=[]
	listbibvirt=[]
	listsemisbn=[]
	nlin=0
	sembib=[]
	bibbinc=[]
	bibcinc=[]
	matriz=[]
	thora1=0
	thora2=0
	thora9=0
	tdisc1=0
	tdisc2=0
	tdisc9=0
	thorasem=0
	sem=999
	try:
# 		fpuds = open(pastaout+codcurso+"PUDs.tex", "w")
# 		fpuds.write('\\newgeometry{top=1in, bottom=2.5in, left=1in, right=1in}\n')
# 		fpuds.write('\\pagestyle{fancy}\n')
# 		fpuds.write('\\lhead{}\n')
		#
		f = open(pastaout+"PUDs"+codcurso+".tex", "w")
		try:
# 			f.write('\\def\\liga{http://www.bn.br/site/pages/servicosProfissionais/agenciaISBN/isbnBusca/FbnBuscaISBNCatalogo.asp?pField=ISBN{\&}pIntPagina=1}\n\n')
			f.write('\\documentclass[12pt,a4paper]{book}\n')
			f.write('\\usepackage{pud2_utf8}\n')
			f.write('\\begin{document}\n')
			f.write('\\thispagestyle{empty}\n')
			localtime = time.asctime( time.localtime(time.time()) )
			f.write('\\newpage\n')
			f.write('\\includegraphics[width=12cm]{pud_logo_ifce.jpg}\n\n')
			f.write('COORDENADORIA DO EIXO DA INDUSTRIA\n\n')
			f.write('\n\n-----------------------------------------\n')
			f.write('\n \n \n \nPUDs DO CURSO\n\n \\textbf{%s}\n\n'%nomecurso.upper())
			f.write('\n\n \\includegraphics[width=10cm]{logo'+codcurso+'.png}\n\n')
			f.write('Atualizado em: %s' % localtime)
			f.write('\n\n-----------------------------------------\n')
			f.write('\n\n \\hyperref[RELATORIO]{PENDÊNCIAS}\n')
			f.write('\n\n\\hyperref[FALTABIBBAS]{Livros básicos não disponíveis na biblioteca}\n')
			f.write('\n\n\\hyperref[FALTABIBCOMP]{Livros complementares não disponíveis na biblioteca}\n')
			f.write('\n\n\\hyperref[PUDSEMBIB]{PUDs sem bibliografia básica ou complementar}\n')
			f.write('\n\n\\hyperref[PUDBIBBASINC]{PUDs com bibliografia básica incompleta}\n')
			f.write('\n\n\\hyperref[PUDBIBCOMPINC]{PUDs com bibliografia complementar incompleta}\n')
			f.write('\n\n\\hyperref[COMPRADOS]{Livros comprados mas não disponíveis na biblioteca}\n')
			f.write('\n\n\\hyperref[BIBVIRT]{Livros que constam na Biblioteca Virtual mas não disponíveis na biblioteca}\n')
			f.write('\n\n-----------------------------------------\n')
			f.write('\n\n \\hyperref[MATRIZ]{MATRIZ CURRICULAR}\n')
			f.write('\n\n \\hyperref[LISTACOMPLETA]{Lista completa de livros}\n')
			
			f.write('\\newpage\n')
			possem=[]
			semant=-1
			nitems=0
			s=[]
			wb = open_workbook(planilha)
			for p in wb.sheets():
				if p.name.find(codcurso)>=0:
					for row in range(p.nrows):
						print p.cell(row,0).value
						v0='999'
						if type(p.cell(row,0).value)==unicode:
							v0=p.cell(row,0).value
						if row>0 and v0.find('999')>=0:
# 							print s
# 							exit()
							disc2=''
							print p.cell(row,1).value
							if p.cell(row,1).value.find('ACE')>=0:
								disc2=p.cell(row,5).value
								folha='ACE'
							elif p.cell(row,1).value.find('EIM')>=0:
								disc2=p.cell(row,5).value
								folha='EIM'
							elif p.cell(row,1).value.find('ECA')>=0:
								disc2=p.cell(row,5).value
								folha='ECA'
	
							if len(disc2)>0:
								for p2 in wb.sheets():
									print p2.name.find(folha)
									if p2.name.find(folha)>=0:
										for row2 in range(p2.nrows):

											v2=p2.cell(row2,2).value
											v0='999'
											if type(p2.cell(row2,0).value)==unicode:
												v0=p2.cell(row2,0).value
											nomedisc2=codif(p2.cell(row2,1).value.strip())
											if row2>0 and (disc2.find(v0)>=0) and (v2>=0) and (v2<999):
												iii=int(p2.cell(row2,2).value) #semestre
												jjj=int(p2.cell(row2,0).value)-900 #cod.disc.
												positx=0
												positxx=-1
												if folha.find('ACE')>=0: # garante ACE aparece em primeiro, depois o curso atual, depois as disciplinas técnicas comuns
													ajuste=1
												else:
													ajuste=0
												#ECA: 000 111 222 333 444
# 												ajuste=0
												print '>>busca: sem=%d cod=%d'%(iii,jjj)
												for x in s: # busca semestre correto
													print 'sem=%d cod=%d'%(x[0],x[3])
													if (x[0])>iii-ajuste and (x[3])>jjj:
														positxx=positx
														break
													positx+=1
												if positxx==-1:
													s.append([iii, p2.name, row2, jjj])													
												else:
													s.insert(positxx,[iii, p2.name, row2, jjj])													
						elif p.cell(row,0).value>0 and p.cell(row,2).value>=0 and p.cell(row,2).value<999 and p.cell(row,3).value>0:
							iii=int(p.cell(row,2).value) #semestre
							jjj=int(p.cell(row,0).value) #cod.disc.
							positx=0
							positxx=-1
							for x in s: # busca semestre correto
								if (x[0])>iii:
									positxx=positx
									break
								positx+=1
							if positxx==-1:
								s.append([iii, p.name, row, jjj])		
							else:
								s.insert(positxx,[iii, p.name, row, jjj])
							nitems+=1
			ss=[]
			for row in s:
			    if int(row[0])!=0:
			        ss.append(row)
			for row in s:
			    if int(row[0])==0:
			        ss.append(row)

			codcursoarq=codcurso
			print 'SS:',ss
			print len(ss)
			print 'FIM'
#   			exit()		
 							
			
			rownum=0
			for reg in ss:
				nlin=0
				arq=''
				folha=codif(reg[1]);
				row=reg[2];
				
				for p in wb.sheets():
					if p.name.find(folha)>=0:		
						if type(p.cell(row,0).value)==float:
							sCod="%d"%p.cell(row,0).value
						else:
							sCod=codif(p.cell(row,0).value.rstrip())
						arq=p.cell(row,1).value.rstrip()	# nao mudar a codificacao
						sNom=codif(p.cell(row,1).value.rstrip())
						
						print sNom
						if sNom.find('(')>=0:
							p1=sNom.find('(')
							sNom=sNom[0:p1-1].strip()
							
						iSem=int(p.cell(row,2).value)
						iCH =int(p.cell(row,3).value)
						iCrd=int(p.cell(row,4).value)
						if type(p.cell(row,5).value)==float:
							sPre="%d"%p.cell(row,5).value
						else:
							sPre=codif(p.cell(row,5).value.rstrip())
						sObj=corrige_pontuacao(codif(p.cell(row,6).value.rstrip()))
						sEme=corrige_pontuacao(codif(p.cell(row,7).value.rstrip()))
						sCon=corrige_pontuacao(codif(p.cell(row,8).value.rstrip()))
						sBB =corrige_pontuacao(codif(p.cell(row,9).value.rstrip()))
						sBC =corrige_pontuacao(codif(p.cell(row,10).value.rstrip()))
						sRes=codif(p.cell(row,11).value.rstrip())
						dDTP=''
						if type(p.cell(row,12).value)==float:
							dDTP=ajustadata(p.cell(row,12).value)
						
						iRev=int(p.cell(row,13).value)
						sRvr=codif(p.cell(row,14).value.rstrip())
						dDTR=''
						if type(p.cell(row,15).value)==float:
							dDTR=ajustadata(p.cell(row,15).value)
						
				if 1==1:#len(iSem)>0:
					# OBS!!!! não pode haver '/' no nome de arquivo!!!!
					arq=arq.replace('\n','') 
					arq=arq.replace('/','') 
					arq=arq.replace(')','') 
					arq=arq.replace('(','') 
					arq=arq.replace('.','') 
					arq=arq.replace(';','') 
					semant=sem
					sem=int(iSem)
					semdesc="S%d"%iSem
					if sem==0:
					   semdesc='Opt'
					semstr="%d"%iSem
					codcursoarq=folha
				
					arq='S%02d' % sem +'-'+codcursoarq+'-PUD-'+arq.replace(' ','')+'.tex'
 					arq=remove_accents(arq)
					
					f2 = open(pastaout+arq, "w")
# 					f2 = codecs.open(pastaout+arq, mode="w", encoding="iso-8859-1")
					try:
						semconteudo=0;
						colnum=0
						
# 						for col in row:
						if 1==1:
# 							if colnum>=6 and colnum<=7 and len(col)>0:
# 								col=col.rstrip()
# 								if col[len(col)-1]==' ':
# 									col[len(col)-1]=='.'
# 								if col[len(col)-1]!='.':
# 									col=col+'.'
# 							col=col.replace(';.','.')
# 							if colnum>=6 and colnum<=9:
# 							    if len(col.strip())==0:
# 							        semconteudo+=1
							semconteudo+=len(sObj)<1
							semconteudo+=len(sEme)<1
							semconteudo+=len(sCon)<1
							semconteudo+=len(sBB)<1
 							semconteudo+=len(sBC)<1
 							
							sCodCompleto=folha+sCod							
							if folha.find('ACE')>=0: 
								sCodCompleto=folha+'0'+sCod[1:3]	# acrescenta o 'E' de ACE						
							
							prerequisitoref=''	
							if len(sPre)>0:
								sPre=' '+sPre
								sPre=sPre.replace(' 2',' ECA2')
								sPre=sPre.replace(' 1',' EIM1')
								sPre=sPre.replace(' 9',' ACE0') # codigo AC9xx agora é ACE0xx
								sPre=sPre[1:]
								prerequisitol=sPre.split(' ')
								for pr in prerequisitol:
									prerequisitoref+=' \\hyperref[PUD:%s]{%s} '%(pr,pr)
									
							if (sBB.count('*')+sBB.count('@')+sBB.count('$')+sBB.count('#'))<3:
								bibbinc.append('S'+semstr+' \\hyperref[PUD:%s]{%s}'%(sCodCompleto.rstrip(),sCodCompleto.rstrip())+' '+sNom+'\n\n')
							if (sBC.count('*')+sBC.count('@')+sBC.count('$')+sBC.count('#'))<1:
								bibcinc.append('S'+semstr+' \\hyperref[PUD:%s]{%s}'%(sCodCompleto.rstrip(),sCodCompleto.rstrip())+' '+sNom+'\n\n')
							if (sBB.count('*')+sBB.count('@')+sBB.count('$')+sBB.count('#')+sBC.count('*')+sBC.count('@')+sBC.count('$')+sBC.count('#'))==0:
								sembib.append('S'+semstr+' \\hyperref[PUD:%s]{%s}'%(sCodCompleto.rstrip(),sCodCompleto.rstrip())+' '+sNom+'\n\n')
								
							if sCon[0:10].find('*')<0:
							    sCon='*'+sCon
							    
							sConr=prepara_tabela(sCon)   
							sBBr=prepara_tabela(sBB)   
							sBCr=prepara_tabela(sBC)   
							sBBr=prepara_bib(sBBr)   
							sBCr=prepara_bib(sBCr)   

							#					
 							list=analisa('(ND-',sBBr,semstr,sCodCompleto,sNom,list)
							listc=analisa('(ND-',sBCr,semstr,sCodCompleto,sNom,listc)
							listcomprados=analisa('(ND-COMPRADO',sBBr,semstr,sCodCompleto,sNom,listcomprados)
							listcomprados=analisa('(ND-COMPRADO',sBCr,semstr,sCodCompleto,sNom,listcomprados)
							listbibvirt=analisa('(ND-BIB.VIRT',sBBr,semstr,sCodCompleto,sNom,listbibvirt)
							listbibvirt=analisa('(ND-BIB.VIRT',sBCr,semstr,sCodCompleto,sNom,listbibvirt)
							
							print sNom
							if sNom.find('Especiais')<0:
								f2.write('\\def\\PUDcurso{%s}\n' % nomecurso)
								if len(sCod)==0:	f2.write('\n\\def\\PUD%s{---}\n' % (header[0]))
								else: 				f2.write('\n\\def\\PUD%s{%s}\n' % (header[0],sCodCompleto))
							
													
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[1],sNom))
							
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[2],semdesc))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[3],iCH))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[4],iCrd))
								if len(prerequisitoref)==0:	f2.write('\n\\def\\PUD%s{---}\n' % (header[5]))
								else:						f2.write('\n\\def\\PUD%s{%s}\n' % (header[5],prerequisitoref))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[6],sObj))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[7],sEme))
								if len(sCon)==0: f2.write('\n\\def\\PUD%s{ & \\\\ & }\n' % (header[8]))
								else: 			 f2.write('\n\\def\\PUD%s{%s}\n' % (header[8],sConr))
							
							
								# remove (ND-xxx)
								if removerND:
									sBBr=sBBr.replace('(ND-ADQUIRIR)','');
									sBBr=sBBr.replace('(ND-COMPRADO)','');
									sBBr=sBBr.replace('(ND-BIB.VIRT)','');
									sBCr=sBCr.replace('(ND-ADQUIRIR)','');
									sBCr=sBCr.replace('(ND-COMPRADO)','');
									sBCr=sBCr.replace('(ND-BIB.VIRT)','');
							
								listcompleta=separalivros(sBBr,semstr,sCodCompleto,sNom,listcompleta)
								listcompleta=separalivros(sBCr,semstr,sCodCompleto,sNom,listcompleta)

#  								sBBr=prepara_biblinkisbn(sBBr)
#  								sBCr=prepara_biblinkisbn(sBCr)
								
								if len(sBB)==0: f2.write('\n\\def\\PUD%s{ & \\\\ & }\n' % (header[9]))
								else:			f2.write('\n\\def\\PUD%s{%s}\n' % (header[9],sBBr))
							
								if len(sBC)==0: f2.write('\n\\def\\PUD%s{ & \\\\ & }\n' % (header[10]))
								else: 			f2.write('\n\\def\\PUD%s{%s}\n' % (header[10],sBCr))
							
							
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[11],sRes))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[12],dDTP))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[13],iRev))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[14],sRvr))
								f2.write('\n\\def\\PUD%s{%s}\n' % (header[15],dDTR))
								
							
							
							
								f2.write('\n\\PUDcorpo\n')
								numpud+=1
								
							if sem!=semant:
								if thorasem>0:
									matriz.append('& & & \\fbox{\\textbf{%3d}} &  \\\\ '%thorasem)
								matriz.append('\hline\n')
								thorasem=0
							if semconteudo==0: # tem conteudo
								cordisciplina=sNom
							elif semconteudo>=4: # não tem conteudo
								cordisciplina='\color{red} '+sNom
							else: # incompleto
								cordisciplina='\color{gray} '+sNom
						
						
							matriz.append(semdesc+'&  \\hyperref[PUD:%s]{%s}'%(sCodCompleto,sCodCompleto)+' & '+cordisciplina+' & %d'%iCH+' & '+prerequisitoref+' \\\\ \n')
							thorasem+=iCH
							if sem>0:
								if int(sCod)<199:
									thora1+=iCH
									tdisc1+=1
								elif int(sCod)<299:
									thora2+=iCH
									tdisc2+=1
								else:
									thora9+=iCH
									tdisc9+=1
							semant=sem
					
							colnum+=1
					finally:
						f2.close()
				#if len
				if len(arq)>0:
					f.write('\\input{%s}\n'%arq)
					f.write('\\newpage\n')
				rownum+=1
			#for row
			####matriz.append('& & & \\fbox{\\textbf{%3d}} &  \\\\ '%thorasem)						
			f.write('\n \\fancyfoot[CO,CE]{Relatório emitido em  %s}' % localtime )


			f.write('\n \n \n')
			f.write('\n\\textbf{%d PUD(s) gerado(s)!!!}\\label{RELATORIO}' % (numpud))
			f.write('\n\n Adquirir %d livro(s) da bib.básica e %d livro(s) da bib.complementar.' % (len(list),len(listc)) )
			f.write('\n\n Falta acrescentar %d livro(s) na bib.básica e %d livro(s) na bib.complementar dos PUDs.' % (len(bibbinc),len(bibcinc)) )
			if len(list)>0:
				f.write('\n\n \\textbf{---> Não constam na biblioteca: %d livro(s) da bibliografia básica:}\\label{FALTABIBBAS}\n' % len(list))
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(list))
				f.write('\\end{enumerate}\n')
			if len(listc)>0:
				f.write('\\newpage\n')
				f.write('\n\n \\textbf{---> Não constam na biblioteca: %d livro(s) da bibliografia complementar:}\\label{FALTABIBCOMP}\n' % len(listc))
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(listc))
				f.write('\\end{enumerate}\n')
			if len(sembib)>0:
				f.write('\\newpage\n')
				f.write('\n\n \\textbf{---> PUD(s) sem bibliografia básica ou sem bibliografia complementar:}\\label{PUDSEMBIB}\n')
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(sembib))
				f.write('\\end{enumerate}\n')
			if len(bibbinc)>0:
				f.write('\\newpage\n')
				f.write('\n\n \\textbf{---> PUD(s) com bibliografia básica incompleta (menos de 3 livros):}\\label{PUDBIBBASINC}\n')
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(bibbinc))
				f.write('\\end{enumerate}\n')
			if len(bibcinc)>0:
				f.write('\\newpage\n')
				f.write('\n\n \\textbf{---> PUD(s) com bibliografia complementar incompleta (nenhum livro):}\\label{PUDBIBCOMPINC}\n')
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(bibcinc))
				f.write('\\end{enumerate}\n')
			if len(listcomprados)>0:
				f.write('\\newpage\n')
				f.write('\n\n \\textbf{---> Comprados mas ainda não constam na biblioteca: %d livro(s):}\\label{COMPRADOS}\n' % len(listcomprados))
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(listcomprados))
				f.write('\\end{enumerate}\n')
			if len(listbibvirt)>0:
				f.write('\\newpage\n')
				f.write('\n\n \\textbf{---> Constam na Biblioteca Virtual mas ainda não constam na biblioteca: %d livro(s):}\\label{BIBVIRT}\n' % len(listbibvirt))
				f.write('\\begin{enumerate}\n \\item ')
				f.write('\n \\item '.join(listbibvirt))
				f.write('\\end{enumerate}\n')
			#
		
			f.write('\\newpage\n')
			f.write('\n\n \n')
			f.write("\\input{Matriz"+codcurso+".tex}\n")
			f.write('\\newpage\n')
			f.write('\n\n \\textbf{---> Lista completa de livros:}\\label{LISTACOMPLETA}\n')
			f.write('\\begin{enumerate}\n \\item ')
			f.write('\n \\item '.join(listcompleta))
			f.write('\\end{enumerate}\n')
			f.write('\\end{document}\n')
			
		finally:
			f.close()
		#	
		try:
			f0 = open(pastaout+"Matriz"+codcurso+".tex", "w")
			f0.write('\n \\begin{longtable}{|c|c|p{9.2cm}|c|p{1.8cm}|}\n')
			f0.write('\n \\multicolumn{5}{c}{\\textbf{MATRIZ CURRICULAR %s}} \\label{MATRIZ} \\\\ \n'%codcurso)
			f0.write('\n \hline \n')
			f0.write('\n \cellcolor{\cinza}Sem & \cellcolor{\cinza}Cod & \cellcolor{\cinza}Disciplina & \cellcolor{\cinza}CH & \cellcolor{\cinza}PR \\\\ \n')
			f0.write('\n \hline \n')
			f0.write('\n'.join(matriz))
			f0.write('\n \hline \n')
			thorasum=thora1+thora2+thora9
			thoraprof=thora1+thora2
			f0.write('\n & &\color{gray} %02d DISC. OBRIGATÓRIAS & \color{gray}%d & \color{gray}%d cr. \\\\ \n'% (tdisc1+tdisc2+tdisc9,(thorasum),(thorasum)/20))
			f0.write('\n & &\color{gray} %02d DISC. PROFISSIONALIZANTES & \color{gray}%d & \color{gray}%d cr. \\\\ \n'% (tdisc1+tdisc2,(thoraprof),(thoraprof)/20))
			f0.write('\n & &\color{gray} %02d DISC. COMUNS/BÁSICAS & \color{gray}%d & \color{gray}%d cr. \\\\ \n'% (tdisc9,(thora9),(thora9)/20))
			f0.write('\n \hline \n')
			f0.write('\n \\end{longtable} \n')
		finally:	
			f0.close()
	except IOError:
		pass

	print '\n%d PUD(s) gerado(s)!!!\n' % (numpud)

 	if 1==1:
	
		subprocess.call(['pdflatex', '-interaction=nonstopmode', pastaout+'PUDs'+codcurso+'.tex'])
 		subprocess.call(['pdflatex', '-interaction=nonstopmode', pastaout+'PUDs'+codcurso+'.tex'])
		subprocess.call(['rm', pastaout+'PUDs'+codcurso+'.aux'])
		subprocess.call(['rm', pastaout+'PUDs'+codcurso+'.log'])
		subprocess.call(['rm', pastaout+'PUDs'+codcurso+'.out'])
# 		subprocess.call(['mv', pastaout+'*.tex','arquivos_tex'])



