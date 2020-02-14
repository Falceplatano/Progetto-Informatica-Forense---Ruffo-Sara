import PySimpleGUI as sg
import os.path
import time
import os
import stat
import win32com.client
from subprocess import call
import webbrowser
import sqlite3
import operator
from collections import OrderedDict
import browserhistory as bh
import glob
import pandas as pd

from os import listdir
from os.path import isfile, join

from datetime import datetime
from datetime import date
import csv

from pandas import ExcelWriter
from xlsxwriter.workbook import Workbook


layout_main = [[ sg.Text('Progetto Ruffo'),],
		  [sg.Button('Proprietà File')], #ok
		  [sg.Button('Visualizza JumpList')],#ok
		  [sg.Button('Visualizza Proprietà file .lnk')],#ok
		  [sg.Button('Proprietà USB Device')],#ok
		  [sg.Button('Analisi Cronologia')],#ok
		  [sg.Button('Shellbags')],#ok
		  [sg.Button('AppCompatCache / Shimcache')],#ok
		  [sg.Button('Application Cache / Appcache')],#ok
		  [sg.Button('Prefetch')], #ok
		  [sg.Button('Superfetch')], #ok
		  [sg.Text('----------------------------------------')],
		  [sg.Button('Unisci CSV')], #ok
		  
		  [sg.Button('Exit')]]

window_main = sg.Window('Main', layout_main)


while True:
	event_proprieta, values_proprieta = window_main.read()
	if event_proprieta == 'Exit':
		break
		
#--------------------------------------------------------------
#Proprietà File


	if event_proprieta == 'Proprietà File':
		proprieta_file = [[sg.Text('Inserisci Path cartella')],
						 [sg.Input(), sg.FolderBrowse()],
						 [sg.Button('Invio'), sg.Button('Exit')]]
		window_proprieta = sg.Window('Proprietà File', proprieta_file)
		
		ev2, vals2 = window_proprieta.read()
		
		now = datetime.now()
		current_time = now.strftime("%H-%M-%S")
		today = date.today()
		today=str(today)
		
		path = vals2[0] + '/'
		
		
		onlyfiles = [f for f in listdir(path) if isfile(join(path, f))]
		
		with open('ProprietaFile' + today + '.csv', 'w', newline='') as file:
			writer = csv.writer(file, delimiter='|')
		   
			for file in onlyfiles:
				access_time = os.path.getatime(path + file)
				local_time = time.ctime(access_time) 
				sg.Print("---Proprietà File---")
				sg.Print("File: ", path + file)
				
				sg.Print("Last modified: %s" % time.ctime(os.path.getmtime(path + file)))
				sg.Print("Created: %s" % time.ctime(os.path.getctime(path + file)))
				sg.Print("Last access time(Local time): ", local_time)
				sg.Print("\n----------------------\n")
				
				writer.writerow([path + file, time.ctime(os.path.getmtime(path + file)), time.ctime(os.path.getctime(path + file)), local_time])
							 
		ev2, vals2 = window_proprieta.read()
		if ev2 == 'Exit':
		   window_proprieta.close()
		   
#-------------------------------------------
#Visualizza JumpList

	event_jumplist, values_jumplist = window_main.read()

	if event_jumplist == 'Visualizza JumpList':
		cwd = os.getcwd()
		os.startfile(cwd + '\\JumpListsView.exe')

		   
#-------------------------------------------
#Proprietà File ink

	event_lnk, values_lnk = window_main.read()

	if event_lnk == 'Visualizza Proprietà file .lnk':
		proprieta_lnk = [[sg.Text('Inserisci Path file .lnk')],
						[sg.Input(), sg.FileBrowse()],
						 [sg.Button('Invio'), sg.Button('Exit')]]
		window_lnk = sg.Window('Proprietà File .lnk', proprieta_lnk)
		
		ev4, vals4 = window_lnk.read()
		
		import win32com.client
		shortcutpath_cur = vals4[0]
		shell = win32com.client.Dispatch('WScript.Shell')
		shortcut = shell.CreateShortcut(shortcutpath_cur)
		target = shortcut.Targetpath
		
		sg.Print('Path target:')
		sg.Print(target)
		
		ev4, vals4 = window_lnk.read()
		if ev4 == 'Exit':
		   window_lnk.close()
		   
#-----------------------------------------------------------
#Proprietà USB Device


	event_usb, values_usb = window_main.read()
	

	if event_usb == 'Proprietà USB Device':
		cwd = os.getcwd()
		os.startfile(cwd + '\\USBDeview.exe')

#-----------------------------------------------------------------------
	
	event_appcompatcache, values_appcompatcache = window_main.read()
	
	if event_appcompatcache == 'AppCompatCache / Shimcache':
		appcompatcache = [[sg.Text('AppCompatCache Parser version 1.4.3.1\nAuthor: Eric Zimmerman\n        c               The ControlSet to parse. Default is to extract all control sets.\n        d               Debug mode\n        f               Full path to SYSTEM hive to process. If this option is not specified, the live Registry will be used\n        t               Sorts last modified timestamps in descending order\n        csv             Directory to save CSV formatted results to. Required\n        csvf            File name to save CSV formatted results to. When present, overrides default name\n\n        dt              The custom date/time format to use when displaying timestamps. See https://goo.gl/CNVq0k for options. Default is: yyyy-MM-dd HH:mm:ss\n        nl              When true, ignore transaction log files for dirty hives. Default is FALSE\n\nExamples: AppCompatCacheParser.exe --csv c:\\temp -t -c 2\n     AppCompatCacheParser.exe --csv c:\\temp --csvf results.csv\n\n         Short options (single letter) are prefixed with a single dash. Long commands are prefixed with two dashes')],
						  [sg.Text('----------------------------------------')],
						  [sg.Text('Inserisci Parametri')],
						  [sg.Input()],
						  
						  [sg.Button('Invio'), sg.Button('Exit')]]
		window_appcompatcache = sg.Window('AppCompatCache / Shimcache', appcompatcache)
		
		ev5, vals5 = window_appcompatcache.read()
		
		os.system("AppCompatCacheParser.exe " + vals5[0])
		
		ev5, vals5 = window_appcompatcache.read()
		if ev5 == 'Exit':
		   window_appcompatcache.close()

#-------------------------------------------------------------------------------------------           
	
	event_jumplist, values_jumplist = window_main.read()
   
	if event_jumplist == 'Shellbags':
		cwd = os.getcwd()
		os.startfile(cwd + '\\ShellBagsExplorer\\ShellBagsExplorer.exe')

#--------------------------------------------------------------------------------------------

	event_cronologiachrome, values_cronologiachrome = window_main.read()
	
	if event_jumplist == 'Analisi Cronologia Chrome':
		cwd = os.getcwd()
		os.startfile(cwd + '\\BrowsingHistoryView.exe')

   
		
		
		#-------------------------------------------------------------------------------
 
	event_applicationcache, values_applicationcache = window_main.read()
	
	if event_applicationcache == 'Application Cache / Appcache':
		cwd = os.getcwd()
		os.startfile(cwd + '\\chromecacheview\\ChromeCacheView.exe')

#-------------------------------------------------------------------------------
	
	event_prefetch, values_prefetch = window_main.read()
	
	if event_prefetch == 'Prefetch':
		cwd = os.getcwd()
		os.startfile(cwd + '\\WinPrefetchView.exe')

#--------------------------------------------------------------------------
	event_superefetch, values_superefetch = window_main.read()
	
	if event_superefetch == 'Superfetch':
		cwd = os.getcwd()
		os.startfile(cwd + '\\SuperFetchTree.exe')
		
#-----------------------------------------------------------------------------

	event_csv, values_csv = window_main.read()
							  		 
	if event_csv == 'Unisci CSV':
		csv_file = [[sg.Text('Inserisci Path cartella contenente csv')],
		[sg.Text('Tutti i csv della cartella verranno uniti in un file exel')],
						 [sg.Input(), sg.FolderBrowse()],					  
						  [sg.Button('Invio'), sg.Button('Exit')]]
		window_csv = sg.Window('Unisci CSV', csv_file)
		
		ev8, vals8 = window_csv.read()
		
		path = vals8[0] + '/'
		today = date.today()
		today=str(today)

		for csvfile in glob.glob(path + '*.csv'):
			workbook = Workbook(csvfile[:-4] + '.xlsx')
			worksheet = workbook.add_worksheet('')
			with open(csvfile, 'rt', encoding='utf8') as f:
				reader = csv.reader(f)
				for r, row in enumerate(reader):
					for c, col in enumerate(row):
						worksheet.write(r, c, col)
			workbook.close()
        
		writer = ExcelWriter(path + today + "combined_CSV.xlsx")
		for filename in glob.glob(path + "*.xlsx"):
			excel_file = pd.ExcelFile(filename)
			(_, f_name) = os.path.split(filename)
			(f_short_name, _) = os.path.splitext(f_name)
			for sheet_name in excel_file.sheet_names:
				df_excel = pd.read_excel(filename, sheet_name=sheet_name)
				df_excel.to_excel(writer, f_short_name, index=False)
				os.remove (filename)

		writer.save()
		
window_main.close()


