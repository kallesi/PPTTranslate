import zipfile
import xml.etree.ElementTree as ET
import openpyxl
import os
from enum import Enum

class Paths(Enum):
  FILE = 0
  FOLDER = 1
  XL = 2
  OUT = 3

class Processor:
  def __init__(self) -> None:
    self._file_path = ""
    self._folder_path = ""
    self._xl_path = ""
    self._out_path = ""

  def use_path(self, file_path: str):
    self._file_path = file_path
    self._folder_path = os.path.dirname(file_path)
    self._xl_path = self._folder_path + r"\Translate.xlsx"
    self._out_path = file_path.replace(".pptx", "_OUT.pptx")
    return self
  
  def get_path(self, path:int):
    if path == Paths.FILE:
      return self._file_path
    if path == Paths.FOLDER:
      return self._folder_path
    if path == Paths.XL:
      return self._xl_path
    if path == Paths.OUT:
      return self._out_path
    
  def extract_text(self):
    # Open the .pptx file as a zip archive
    z = zipfile.ZipFile(self._file_path, "r")
    # Get a list of all slide files
    slide_files = [name for name in z.namelist() if "ppt/slides/slide" in name and not "-rels" in name]
    # Create a new Excel workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    # Loop over each slide
    for slide_file in slide_files:
      # Read the XML data
      slide_xml = z.read(slide_file)
      # Parse the xml data
      root = ET.fromstring(slide_xml)
      # Define the namespaces
      namespaces = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
      # Loop over each text elem
      for elem in root.findall(".//a:t", namespaces):
        ws.append([elem.text])
      wb.save(self._xl_path)
    z.close()
    
  def replace_text(self):
    zin = zipfile.ZipFile(self._file_path, "r")
    zout = zipfile.ZipFile(self._out_path, "w")
    slide_files = [name for name in zin.namelist() if "ppt/slides/slide" in name and not "-rels" in name]
    for item in zin.infolist():
      if item.filename not in slide_files:
        zout.writestr(item, zin.read(item.filename))
    wb = openpyxl.load_workbook(self._xl_path)
    ws = wb.active
    translations = [row[1].value for row in ws.iter_rows()]
    # Loop over each file
    for slide_file in slide_files:
      slide_xml = zin.read(slide_file)
      # Parse xml data
      root = ET.fromstring(slide_xml)
      namespaces = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
                    "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
      for i, elem in enumerate(root.findall(".//a:t", namespaces)):
        if translations[i] == None:
          elem.text = ""
        else:
          elem.text = str(translations[i])
      # Write modified xml back to slide
      slide_xml_new = ET.tostring(root).decode()
      zout.writestr(slide_file, slide_xml_new)
    zin.close()
    zout.close()
      