import sys

class Candidate:	
	def __init__(self,name,email,subject,template,tech,status):
		self.name=name
		self.email=email
		self.subject=subject
		self.template=template
		self.tech = tech
		self.status = status


def sendEmail(to, subject: str, template: str):
	"""Send the email to the candidates

	It sends an email using mariano.selvaggi@whiteprompt.com with the right template
	"""
	import smtplib, string

	gmail_user = 'mariano.selvaggi@whiteprompt.com'
	gmail_pass = "M4r14n0."
	
	Body = "\n".join(["From: %s" % gmail_user,"To: %s" % to,"Subject: %s" % subject, template])

	try:
		server = smtplib.SMTP('smtp.gmail.com')#,587)
		#server.ehlo()
		server.starttls()	
		server.login(gmail_user,gmail_pass)				
		server.sendmail(gmail_user, to, Body)
		#server.close()
		server.quit()
	except Exception as e:		
		raise e


def getTemplate(candidate):
	"""Get the right template to include in the body of the email

	Get the right file and parse the information in order to create a good speech
	"""
	template=""	
	try:		
		with open("templates/" + candidate.template + ".txt", "r") as infile:
			template = infile.read()
			template = template.replace("[Name]",candidate.name)
			template = template.replace("[Tech]",candidate.tech)
	except FileNotFoundError:		
		raise Exception('There is no file for this template')
	except Exception as e:		
		raise e
	return template


def getCandidatesFromTxt(file):
	"""Get the candidates from source

	It gets the different candidates from the txt file with all the information needed
	"""
	candidates=[]
	try:
		with open(file, "r") as infile:			
			for line in infile:				
				items=line.split('|')
				status = ""				
				if len(items) > 5:
					status=str(items[5])
				candidates.append(Candidate(items[0].strip(),items[1].strip(),items[2].strip(),items[3].strip(),items[4].strip(),status))
	except FileNotFoundError:		
		raise Exception('There is no file for this txt')
	except Exception as e:				
		raise e
	
	return candidates


def getCandidatesFromExcel(file):
	"""Get the candidates from source

	It gets the different candidates from the excel file with all the information needed
	"""
	from xlrd import open_workbook

	candidates	= []
	
	try:
		wb = open_workbook(file)

		sheet = wb.sheets()[0]
		number_of_rows = sheet.nrows
		number_of_columns = sheet.ncols
		
		for row in range(1,number_of_rows):		
			values = []
			for col in range(0,number_of_columns):
				value = (sheet.cell(row,col).value)			
				try:
					value = str(value)
				except:
					value = ""
				finally:
					values.append(value)
			
			candidate = Candidate(name=values[0],email=values[1],subject=values[2],template=values[3],tech=values[4],status=values[5])
			candidates.append(candidate)
	except FileNotFoundError:
		raise Exception("There is no excel file")
	except Exception as e:
		raise e
	finally:
		return candidates

def markFileAsSent(file,candidate,row):
	"""Mark the row as sent

	It marks the specific email in the txt file
	"""			
	if file.endswith(".txt"):
		try:
			totalline = ""
			#read the file and create changes
			with open(file, "r") as infile:
				j = 1				
				for line in infile:
					newline=line.rstrip()
					if j == row:					
						newline = newline + " | sent"
					j = j + 1
					totalline = totalline + newline + "\n"
			#write the file with changes
			with open(file,"w") as infile:
				infile.write(totalline.rstrip())
		except FileNotFoundError:		
			raise Exception('There is no file for this txt')
		except Exception as e:
			raise e
	else:		
		try:
			from xlrd import open_workbook
			from xlutils.copy import copy			
	
			rb = open_workbook(file)
			wb = copy(rb)

			sheet = wb.get_sheet(0)
			sheet.write(row,5,"sent")

			wb.save(file)			
		except FileNotFoundError:
			raise Exception('There is no file for this txt')			
		except Exception as e:
			raise e

def main(file):
	import smtplib, string

	candidates = []
	i=1

	#Get the list of candidates from files
	try:
		if file.endswith(".txt"):
			candidates =  getCandidatesFromTxt(file)			
		elif file.endswith(".xlsx") or file.endswith(".xls"):
			candidates =  getCandidatesFromExcel(file)			
		else:
			print("you must input either a txt or an excel file")	
	except:
		print(sys.exc_info()[0])

	#looping the candidates to send each email
	try:
		server = smtplib.SMTP('smtp.gmail.com')#,587)
		#server.ehlo()
		server.starttls()		
		server.login(gmail_user,gmail_pass)		
		for candidate in candidates:
			try:
				template = getTemplate(candidate) #get the text to sent the email
				if "sent" not in candidate.status:
					body = "\n".join(["From: %s" % gmail_from,"To: %s" % candidate.email,"Subject: %s" % candidate.subject, candidate.template])					
					server.sendmail(gmail_user, candidate.email, body)
					markFileAsSent(file,candidate,i) #mark the file so the next time the same email is not sent again
					print("new email to:" + candidate.email + " using %s" % candidate.template + "\n")
				i=i+1
			except Exception as e:
				print("Unexpected error:", e)
				break
		server.quit()
		#server.close()
	except Exception as ex:
		print("Error in smtp:", ex)

def readConfig():
	"""Read the config file

	Obtain the most important key to make execute the program
	"""
	import json
	data = []
	with open('config.json') as json_data_file:
		data =json.load(json_data_file)
	return data

 #get the config settings
data = readConfig()
gmail_user=data["mail"]["user"]
gmail_pass=data["mail"]["pass"]
gmail_from=data["mail"]["from"]

#start the process
main(input("Please enter the file name (include the file ext): "))