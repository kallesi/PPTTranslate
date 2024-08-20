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

To get Google Translate functionality:

`pip install playwright`

This implementation doesn't require binaries, but does require a chrome version. You can specify your chrome version in `constants.py`. I use Brave browser, but it should work with any chromium based browser. 

Simply clone this repository or download it in a zip. 
Run `main.py --help` to get started. 

# Example Use

I want to translate a powerpoint file to Japanese:

`cd {project directory}`

Extract the text into `Translate.xlsx` - then auto translate via Google Translate.

`main.py extract {pptx path} --auto --source en --to ja`

Make necessary adjustments. The `Translate.xlsx` will be in the same directory as your original powerpoint. 

`main.py merge {pptx path}`

Your translated text will be merged into a new pptx file. 

