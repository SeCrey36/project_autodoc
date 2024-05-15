import os
import urllib.request

dirname = 'C:/Apache24/htdocs/autodoc/files'
dirfiles = os.listdir(dirname)
print(dirfiles)

url = 'http://127.0.0.1/autodoc/files/'+dirfiles[0]
urllib.request.urlretrieve(url, f'./files/{dirfiles[0]}')

file_path = f'C:/Apache24/htdocs/autodoc/files/{dirfiles[0]}'
os.unlink(file_path)