import win32file
import time
import pywintypes
import os
import subprocess
from findNamedPipe import getNamedPipes

def sendCommandsToGME():
	print("Sending Commands to GME")
	fileHandles = []
	namedPipes = []
	try:
		namedPipes = getNamedPipes("tasck_behavior_library.mga")
	except pywintypes.error:
		pass

	for pipeFile in namedPipes:
		print(pipeFile)
		handle = win32file.CreateFile(pipeFile, win32file.GENERIC_READ | win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0, None)
		fileHandles.append(handle)

	for handle in fileHandles:
		try:
			win32file.WriteFile(handle, "Close")
		except pywintypes.error as e:
			print(str(e))
			handle.Close()
			fileHandles.remove(handle)

	for handle in fileHandles:
		try:
			win32file.ReadFile(handle, 256)
		except pywintypes.error as e:
			print(str(e))
			handle.Close()
			fileHandles.remove(handle)

	for handle in fileHandles:
		try:
			win32file.WriteFile(handle, "Open")
		except pywintypes.error as e:
			print(str(e))
			handle.Close()
			fileHandles.remove(handle)

	for handle in fileHandles:
		handle.Close()

	return True

confirmationBox = os.path.dirname(os.path.realpath(__file__)) + r"\ConfirmationBox.exe"
message  = "Do you want to save your Open Mga File?\n"
message += "Doing so will cuase you to loose your undo\n"
message += "history, because the file will get closed"
label    = "Mga File Save Confirmation"
result   = subprocess.call([confirmationBox, message, label, "YESNOCANCEL"])
result == 2 and exit(-1)
result == 6 and sendCommandsToGME()


# for pipeFile in getNamedPipes("tasck_behavior_library.mga"):
# 	fileHandle = win32file.CreateFile(pipeFile, win32file.GENERIC_READ | win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0, None)
# 	win32file.ReadFile(fileHandle, 256);
# 	fileHandle.Close();

# for pipeFile in getNamedPipes("tasck_behavior_library.mga"):
# 	fileHandle = win32file.CreateFile(pipeFile, win32file.GENERIC_READ | win32file.GENERIC_WRITE, 0, None, win32file.OPEN_EXISTING, 0, None)
# 	win32file.WriteFile(fileHandle, "open");
# 	fileHandle.Close();

# i = 10 
# while (i):
# 	i = i - 1
# 	left, data = win32file.ReadFile(fileHandle, 4096)
# 	print data # prints \rprint "hello"

# i = 10
# while (i):
# 	i = i - 1
# 	win32file.WriteFile(fileHandle, "hello back\n")