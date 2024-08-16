# PPT Translate

A tool written in Python 3 for a very common office task that takes up way too much time. 

### Usecase

1. You have a xml PowerPoint file (ends in .pptx) which you have to translate into another language. 

2. With this tool, you can easily copy-paste all the PowerPoint content into Google Translate, and get a fully formatted PowerPoint in another language ready for distribution (after some proofreading)


### The Tool

This is not production code. Just a demo of the capabilities of XML parsing and the relative ease of doing this via Python open XML vs doing it via the Office COM Object Model in VBA. 

Also, this hasn't been tested on Mac/Linux although I don't see why it wouldn't work on those platforms. 

To get started, you need a recent version Python 3 and a few libraries. 

`pip install zipfile openpyxl` 

Simply clone this repository or download it in a zip. `main.py` will contain some general instructions on how to get started with the tool. 
