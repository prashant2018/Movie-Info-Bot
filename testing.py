# Movie Assistor - This script scans the directory for movies and returns the ratings and riviews in a spreadsheet.

import os
import guessit
import threading
import pyperclip
import sys
import time
import getInfo
# List of all possible video formats
name_list = ["webm","mkv","flv","vob","ogv","ogg","drc","gif","gifv","mng","avi","moc","qt","wmv","yuv","rm","rmvb","asf",
			 "mp4","m4p","m4v","mpg","mp2","mpe","mpeg","m2v","svi","3gp","3g2","mxf","roq","nsv","f4p","f4a","f4b"]

#list will contain movie name_list
movie_names = []

#list will contain movie information
movie_info = [{},{},{},{}]

# path of directory is stored here
dir_path = ""

# Input of directory path
def input_directory():

	if len(sys.argv) == 1:
		print("Copy folder's path and paste here : ")
		dir_path = input()
	else:
		dir_path = pyperclip.paste()

	if os.path.exists(dir_path):
		os.chdir(dir_path)

	else:
		print("Directory Does Not Exists")
		return 999

# movie name extracting
def movie_title(item):
	title = ""
	if "sample" in item.lower():
		return ""
	try:
		data = guessit.guessit(item)
	except:
		return ""

	try:
		extension = data['container']
	except:
		return ""

	if extension in name_list:
		try:
			title = data['title']
		except:
			title = ""
	return title

# Walkthrough all movie folders
def walk_dir():
	title=""

	for	folderName,	subfolders,	filenames in os.walk(os.getcwd()):
		for	filename in	filenames:
			title = movie_title(filename)
			if title != "":
				movie_names.append(title)

#finally calling and displaying for testing purpose
def func_call():
	input_directory()
	#print(dir_path)
	walk_dir()

	for name in movie_names:
		print(name)

	download_review()

def thread_download(num):
	#download information
	temp_list = []
	temp = num
	check = getInfo.getInfo(movie_names[num])
	try:
		temp_list.append(check)
		if check == "":
			movie_info[temp] = temp_list
			return movie_info[temp]
		else:
			num = num + 4
	except:
		return movie_info[temp]


def download_review():
	thread_list = list()
	mlength = len(movie_names)
	buff = mlength//4
	n = [0,1,2,3]
	for i in range(0,4):
		temp_thread = threading.Thread(target = thread_download, args=[ n[i] ])
		thread_list.append(temp_thread)
		temp_thread.start()
	for threadObj in thread_list:
		threadObj.join()

def display():
	for i in movie_info:
		print(i)

func_call()
display()
