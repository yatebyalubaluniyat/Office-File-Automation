#!/usr/bin/python

"""                                                      #
    ####    ####   ####   ##   ####  ####        #      ##
    ####    ####   #### ##     ####  ####       ###   ##### 
  #################### # ######### ## #####    ##### ######
   ####    ####   ####  ##    #### ## ####     ##### #####
   ####    ####   ####    ##  ####    ####      ###  ###
                                                 #   ##
# ENJOY!!                                            #
"""
#############################################################
## Description: Opens Office Document files inside temp folder
#############################################################

import os, shutil, sys, time
import win32com.client

print "#"*50
print "Copyrights H A K U N A   M A T A T A :) Enjoy!!"
print "#"*50
print " "

print "Description: This application checks if the executed office document executes calc.exe"
print " "

path = os.getcwd()

filesprev = ""
counter = 0
counter1 = 0
fileflag = 0

flg = os.popen("tasklist | findstr /I \"xampp\"").read()
# print "flg = "+str(flg)
if flg == "":
	print "WARNING!!! Your xampp server is not running.. Please start it now.. I am waiting..."
	raw_input("Press enter to continue")
	flg1 = os.popen("tasklist | findstr \"xampp-control.exe\"").read()
	if flg1 == "":
		print "Xampp is still not running... "
		raw_input("press enter to exit... ")
		sys.exit(0)
		
filetype = raw_input("Enter the document type( eg. .doc .ppt .xls etc): ")
print " "
print "Please put all {} files and folders containing {} inside \"temp\" folder".format(str(filetype), str(filetype))
raw_input("If you are ready, press enter to continue")

md = raw_input("Do you want to 1.Move 2.Delete checked files (enter 1 or 2)")
if str(md) == "1": 
	try:
		os.mkdir("MovedFiles")
	except:
		print "MovedFiles folder exists.... please rename/delete it to continue"
		raw_input("If you are done, press anter to continue")
		os.mkdir("MovedFiles")
elif str(md) != "2":
	print "Wrong input.. proceeding with Move option"
	md = "1"

##### Deleting files starting with "~" #####

for folderscheck in os.listdir("temp"):							#temp dir
	fileName0, fileExtension0 = os.path.splitext(folderscheck)
	if folderscheck.startswith("~"):
		os.remove(folderscheck)
		
	foldercheckpath = path+ "\\temp\\" + folderscheck			#temp dir
	if os.path.isdir(path+ "\\temp\\"+folderscheck):
		for filescheck in os.listdir(foldercheckpath):
			if filescheck.startswith("~"):
				try:
					os.remove(filescheck)
				except:
					pass

############################################

folderpath1 = path+ "\\temp\\"
			
for folders in os.listdir("temp"):								#temp dir
	fileName1, fileExtension1 = os.path.splitext(folders)
	if folders.startswith("~"):
		os.remove(folders)
	# if fileExtension1 != "":
	if fileExtension1 == str(filetype):
		
		####################
		#### POWERPOINT ####
		if str(filetype) == ".ppt" or str(filetype) == ".pptx":
			PptApplication = win32com.client.Dispatch("PowerPoint.Application")
			try:
				PptApplication.Visible = True
			except:
				time.sleep(10)
				PptApplication = win32com.client.Dispatch("PowerPoint.Application")
				PptApplication.Visible = True
			# PptApplication.Visible = True
			fopen = open("currentFile.txt", 'wb')
			fopen.write(fileName1)
			fopen.close()
			try:
				ppt = PptApplication.Presentations.Open(folderpath1+"\\"+folders)
				print "Checking " + folders
			# raw_input("please exit now")
				
			except Exception as e:
				print e
				# raw_input("please exit now")
				
			try:
				PptApplication.Quit()
			except Exception as e1:
				print e1
				
			flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
			if flag != "":
				print "Calc executed in file : "+ folders
				# print "flag :", flag
				os.system("pause")
			##############################################
			## Moving Files to MovedFiles folder #########
			##############################################
			if str(md) == "1":
				try:
					if counter1 != 0:
						src = folderpath1 + "\\" + filesprev1
						# print src
						# os.system("pause")
						dst = path+"\\MovedFiles"
						shutil.move(src, dst)
						
				except Exception as e2:
					print e2
					pass
				counter1 += 1
				filesprev1 = folders
			########################
			## Deleting Files ######
			########################
			elif str(md) == "2":
				if counter1 != 0:
					src = folderpath1 + "\\" + filesprev1
					os.remove(src)
				counter1 += 1
				filesprev1 = folders
			
		###############
		#### EXCEL ####
		elif str(filetype) == ".xls" or str(filetype) == ".xlsx":
			ExcelApplication = win32com.client.Dispatch("Excel.Application")
			try:
				ExcelApplication.Visible = False
			except:
				# WordApplication.Visible = True
				time.sleep(10)
				ExcelApplication = win32com.client.Dispatch("Excel.Application")
				ExcelApplication.Visible = False
			# ExcelApplication.Visible = False
			fopen = open("currentFile.txt", 'wb')
			fopen.write(fileName1)
			fopen.close()
			try:
				excel = ExcelApplication.Workbooks.Open(folderpath1+"\\"+folders)
				print "Checking " + folders
			# raw_input("please exit now")
				
			except Exception as e:
				print e
				# raw_input("please exit now")
				
			try:
				ExcelApplication.Quit()
			except Exception as e1:
				print e1
				
			flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
			if flag != "":
				print "Calc executed in file : "+ folders
				# print "flag :", flag
				os.system("pause")
			##############################################
			## Moving Files to MovedFiles folder #########
			##############################################
			if str(md) == "1":
				try:
					if counter1 != 0:
						src = folderpath1 + "\\" + filesprev1
						# print src
						# os.system("pause")
						dst = path+"\\MovedFiles"
						shutil.move(src, dst)
						
				except Exception as e2:
					print e2
					pass
				counter1 += 1
				filesprev1 = folders
			########################
			## Deleting Files ######
			########################
			elif str(md) == "2":
				if counter1 != 0:
					src = folderpath1 + "\\" + filesprev1
					os.remove(src)
				counter1 += 1
				filesprev1 = folders
			
		##############
		#### WORD ####
		elif str(filetype) == ".doc" or str(filetype) == ".docx":
			WordApplication = win32com.client.Dispatch("Word.Application")
			try:
				WordApplication.Visible = False
			except:
				# WordApplication.Visible = True
				time.sleep(10)
				WordApplication = win32com.client.Dispatch("Word.Application")
				WordApplication.Visible = False
			fopen = open("currentFile.txt", 'wb')
			fopen.write(fileName1)
			fopen.close()
			try:
				word = WordApplication.Documents.Open(folderpath1+"\\"+folders)
				print "Checking " + folders
			# raw_input("please exit now")
				
			except Exception as e:
				print e
				# raw_input("please exit now")
				
			try:
				WordApplication.Quit()
			except Exception as e1:
				print e1
				
			flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
			if flag != "":
				print "Calc executed in file : "+ folders
				# print "flag :", flag
				os.system("pause")
			##############################################
			## Moving Files to MovedFiles folder #########
			##############################################
			if str(md) == "1":
				try:
					if counter1 != 0:
						src = folderpath1 + "\\" + filesprev1
						# print src
						# os.system("pause")
						dst = path+"\\MovedFiles"
						shutil.move(src, dst)
						
				except Exception as e2:
					print e2
					pass
				counter1 += 1
				filesprev1 = folders
			########################
			## Deleting Files ######
			########################
			elif str(md) == "2":
				if counter1 != 0:
					src = folderpath1 + "\\" + filesprev1
					os.remove(src)
				counter1 += 1
				filesprev1 = folders
				
				
	counter = 0			
	folderpath = path+ "\\temp\\" + folders
	if os.path.isdir(folderpath):
		for files in os.listdir(folderpath):
			if files.startswith("~"):
				fileflag = 1
				# os.remove(files)
			fileName, fileExtension = os.path.splitext(files)
			
			# if fileExtension != "":
			if fileExtension == str(filetype) and fileflag == 0:
				
				####################
				#### POWERPOINT ####
				if str(filetype) == ".ppt" or str(filetype) == ".pptx":
					PptApplication = win32com.client.Dispatch("PowerPoint.Application")
					try:
						PptApplication.Visible = True
					except:
						time.sleep(10)
						PptApplication = win32com.client.Dispatch("PowerPoint.Application")
						PptApplication.Visible = True
					# PptApplication.Visible = True
					fopen = open("currentFile.txt", 'wb')
					fopen.write(fileName)
					fopen.close()
					try:
						ppt = PptApplication.Presentations.Open(folderpath+"\\"+files)
						print "Checking " + files
					# raw_input("please exit now")
						
					except Exception as e:
						print e
						# raw_input("please exit now")
						
					try:
						PptApplication.Quit()
					except Exception as e1:
						print e1
						
					flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
					if flag != "":
						print "Calc executed in file : "+ files
						# print "flag :", flag
						os.system("pause")
					##############################################
					## Moving Files to MovedFiles folder #########
					##############################################
					if str(md) == "1":
						try:
							if counter != 0:
								src = folderpath + "\\" + filesprev
								# print src
								# os.system("pause")
								dst = path+"\\MovedFiles"
								shutil.move(src, dst)
								
						except Exception as e2:
							print e2
							pass
						counter += 1
						filesprev = files
					########################
					## Deleting Files ######
					########################
					elif str(md) == "2":
						if counter != 0:
							src = folderpath + "\\" + filesprev
							os.remove(src)
						counter += 1
						filesprev = files
					
				###############
				#### EXCEL ####
				elif str(filetype) == ".xls" or str(filetype) == ".xlsx":
					ExcelApplication = win32com.client.Dispatch("Excel.Application")
					try:
						ExcelApplication.Visible = False
					except:
						time.sleep(10)
						ExcelApplication = win32com.client.Dispatch("Excel.Application")
						ExcelApplication.Visible = False
					# ExcelApplication.Visible = False
					fopen = open("currentFile.txt", 'wb')
					fopen.write(fileName)
					fopen.close()
					try:
						excel = ExcelApplication.Workbooks.Open(folderpath+"\\"+files)
						print "Checking " + files
					# raw_input("please exit now")
						
					except Exception as e:
						print e
						# raw_input("please exit now")
						
					try:
						ExcelApplication.Quit()
					except Exception as e1:
						print e1
						
					flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
					if flag != "":
						print "Calc executed in file : "+ files
						# print "flag :", flag
						os.system("pause")
					##############################################
					## Moving Files to MovedFiles folder #########
					##############################################
					if str(md) == "1":
						try:
							if counter != 0:
								src = folderpath + "\\" + filesprev
								# print src
								# os.system("pause")
								dst = path+"\\MovedFiles"
								shutil.move(src, dst)
								
						except Exception as e2:
							print e2
							pass
						counter += 1
						filesprev = files
					########################
					## Deleting Files ######
					########################
					elif str(md) == "2":
						if counter != 0:
							src = folderpath + "\\" + filesprev
							os.remove(src)
						counter += 1
						filesprev = files
					
				##############
				#### WORD ####
				elif str(filetype) == ".doc" or str(filetype) == ".docx":
					WordApplication = win32com.client.Dispatch("Word.Application")
					try:
						WordApplication.Visible = False
					except:
						# WordApplication.Visible = True
						time.sleep(10)
						WordApplication = win32com.client.Dispatch("Word.Application")
						WordApplication.Visible = False
					print "Checking " + files
					fopen = open("currentFile.txt", 'wb')
					fopen.write(fileName)
					fopen.close()
					try:
						# os.chdir(folderpath)
						# print os.getcwd()
						word = WordApplication.Documents.Open(folderpath+"\\"+files)
						print "Checking " + files
						# os.chdir(path)
					# raw_input("please exit now")
						
					except Exception as e:
						print e
						# raw_input("please exit now")
						
					try:
						WordApplication.Quit()
					except Exception as e1:
						print e1
						
					flag = os.popen("tasklist | findstr /I \"calc.exe\"").read()
					if flag != "":
						print "Calc executed in file : "+ files
						# print "flag :", flag
						os.system("pause")
					##############################################
					## Moving Files to MovedFiles folder #########
					##############################################
					if str(md) == "1":
						try:
							if counter != 0:
								src = folderpath + "\\" + filesprev
								# print src
								# os.system("pause")
								dst = path+"\\MovedFiles"
								shutil.move(src, dst)
								
						except Exception as e2:
							print e2
							pass
						counter += 1
						filesprev = files
					########################
					## Deleting Files ######
					########################
					elif str(md) == "2":
						if counter != 0:
							src = folderpath + "\\" + filesprev
							os.remove(src)
						counter += 1
						filesprev = files
			fileflag = 0			
				
		
print " "
print "#"*50
print "Copyrights H A K U N A   M A T A T A :) Enjoy!!"
print "#"*50
print " "
raw_input("Finished processing... Press any key to Exit")