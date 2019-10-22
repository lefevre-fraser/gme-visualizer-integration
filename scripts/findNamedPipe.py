import os
import win32file
def getNamedPipes(pipeFile):
	file_list = []
	for file in win32file.FindFilesW("\\\\.\\pipe\\*{}".format(pipeFile)):
		if file[8].endswith("\\{}".format(pipeFile)):
			file_list.append(os.path.join("\\\\.\\pipe\\", file[8]))
	return file_list

# for pipe in getNamedPipes("myPET.mga"):
# 	print(pipe)