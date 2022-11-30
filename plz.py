import re
from urllib import request
import wget
import os
import docx2txt 
import zipfile
import subprocess
homepath = ""
 

def downloadzip(url, name):
  path = homepath + "/zips/" + name + "/"
  docpath = homepath + "/Docs/" + name + "/"
  print(path)
  if (not os.path.isdir(path)):
    os.mkdir(path)
  if (not os.path.isdir(docpath)):
    os.mkdir(docpath)

  html=request.urlopen(url) 
  html_contents=str(html.read().decode("cp949"))
  series_list = re.findall(r"(https)(.+)(.zip\">)", html_contents) 
  for url in series_list: 
    tmp_url="".join(url) 
    final_url = tmp_url[:tmp_url.find('"')]
    while True:
      try:
        zipname = wget.download(final_url, out=path)
        unzip(zipname, name);
        break
      except:
        continue


def unzip(filename, name):

  path = homepath + "/Docs/" + name + "/"
  try:
    with zipfile.ZipFile(filename, 'r') as zipObj:
     listOfFileNames = zipObj.namelist()
     for tmpfile in listOfFileNames:
        if tmpfile.endswith("doc") or tmpfile.endswith("docx"):
           zipObj.extract(tmpfile, path=path)
           subprocess.call(['soffice', '--headless', '--convert-to', 'txt:Text', path + tmpfile,'-outdir', homepath + "/txts/" + name + "/"])
        if tmpfile.endswith("zip"):
          zipObj.extract(tmpfile)
          with zipfile.ZipFile(tmpfile, 'r') as test:
            tmplist = test.namelist()
            for tmp in tmplist:
              if tmp.endswith("doc") or tmp.endswith("docx"):
                test.extract(tmp, path=path)
                subprocess.call(['soffice', '--headless', '--convert-to', 'txt:Text', path + tmp,'-outdir', homepath + "/txts/" + name + "/"]) 
  except:
    f = open("log.txt", 'a',encoding='utf-8')
    f.write(name + " : unzip error" + "\n")
    f.close()

   



if __name__ == '__main__':
  homepath = os.getcwd()
  if (not os.path.isdir(homepath + "/Docs")):
    os.mkdir(homepath + "/Docs")
  if (not os.path.isdir(homepath + "/txts")):
    os.mkdir(homepath + "/txts")
  if (not os.path.isdir(homepath + "/zips")):
    os.mkdir(homepath + "/zips")
  url="https://www.3gpp.org/ftp/tsg_sa/WG3_Security/"
  html=request.urlopen(url) 
  html_contents=str(html.read().decode("cp949"))
  series_list = re.findall(r"(https)(.+)(>TSGS3)", html_contents) 

  for url in series_list: 
    tmp_url="".join(url) 
    final_url = tmp_url[:tmp_url.find('"')] + "/Docs"

    try:
      version = int(re.findall(r'\d+',final_url[final_url.find("TSGS3_") + len("TSGS3_"):])[0])
      if not ((108 <= version and version <= 109) or (74 <= version and version <= 76)):
        continue
    except:
      break
    print("Downloading : " + final_url)

    name = final_url[final_url.find("TSGS3_"):final_url.find("/Docs")]
    downloadzip(final_url, name)
