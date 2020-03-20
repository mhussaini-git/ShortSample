#!/usr/bin/python3
# This uses win32com to automate the comparison of two Microsoft Word files.
# Make sure to have win32com installed for your environment and python version:

from os import getcwd, path
from sys import argv, exit
from win32com import client #makesure you check the right client

def die(message):
    print (message)
    exit(1)

def cmp(original_file, modified_file):
  dir = getcwd() + '\\'
  print('Working...')

  # some file checks
  if not path.exists(dir+original_file):
    die('Original file does not exist')
  if not path.exists(dir+modified_file):
    die('Modified file does not exist')
  cmp_file = dir + original_file[:-5]+'_cmp_'+modified_file # use input filenames, but strip extension
  if path.exists(cmp_file):
    die('Comparison file already exists... aborting\nRemove or rename '+cmp_file)

  # actual Word automation
  app = client.gencache.EnsureDispatch("Word.Application")
  app.CompareDocuments(app.Documents.Open(dir + original_file), app.Documents.Open(dir + modified_file))
  app.ActiveDocument.ActiveWindow.View.Type = 3 # prevent that word opens itself
  app.ActiveDocument.SaveAs(cmp_file)

  print('Saved comparison as: '+cmp_file)
  app.Quit()

def main():
  if len(argv) != 3:
    die('Usage: wrd_cmp <original_file> <modified_file>')
  cmp(argv[1], argv[2])

if __name__ == '__main__':
  main()