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
## Description: Opens ppt files in folders inside temp folder
#############################################################

import os, shutil
import win32com.client

path = os.getcwd()
os.mkdir("MovedFiles")
filesprev = ""
counter = 0

for folderscheck in os.listdir("temp"):							#temp dir
	foldercheckpath = path+ "\\temp\\" + folderscheck
	for filescheck in os.listdir(foldercheckpath):
		if filescheck.startswith("~"):
			os.remove(filescheck)
			
for folders in os.listdir("temp"):								#temp dir
	folderpath = path+ "\\temp\\" + folders
	# print folderpath
	# os.system("pause")
	for files in os.listdir(folderpath):
		Application = win32com.client.Dispatch("PowerPoint.Application") 
		Application.Visible = True
		fileName, fileExtension = os.path.splitext(files)
		#### Writing log ####
		fopen = open("currentPPT.txt", 'wb')
		fopen.write(fileName)
		fopen.close()
		#####################
		if files.endswith(".ppt"):
			print files
			try:
				ppt = Application.Presentations.Open(folderpath+"\\"+files)
				# raw_input("please exit now")
					
			except Exception as e:
				print e
				# raw_input("please exit now")
		try:
			Application.Quit()
		except Exception as e1:
			print e1
			flag = os.system("tasklist | findstr \"calc.exe\"")
			if flag == 0:
				print "Calc executed in file : "+files
				# print "flag :", flag
				os.system("pause")
				
		##############################################
		## Moving Files to MovedFiles folder #########
		##############################################
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
		
print "exittingg................."
os.system("pause")