from Processor import Processor
from Processor import Paths
from constants import *
import argparse

def main():
  parser = argparse.ArgumentParser(prog="PPT Translate",
                                   description="Translate your powerpoints easily",
                                   epilog="https://www.github.com/kallesi")
  
  parser.add_argument("task", choices=["extract", "merge"], type=str,
                      help="`extract` to extract text for translation, `merge` to replace text in PowerPoint")
  parser.add_argument("path", metavar=r"C:\PowerPoint\Path.pptx", type=str,
                      help="Specify full path of your .pptx file")
  parser.add_argument("-a", "--auto", required=False, action="store_true",
                      help="Auto translate by Google Translate using Playwright")
  parser.add_argument("-s", "--source", required=False, default="en", type=str,
                      help="From language - source language. Defaults to `en` for English")
  parser.add_argument("-t", "--to", required=False, default="zh-CN", type=str,
                      help="To language - destination language. Defaults to `zh-CN` for simplified Chinese")

  args = parser.parse_args()
  ppt_path = args.path
  task = args.task
  if task == "extract":
    if args.auto == True:
      Processor().set_path(ppt_path).extract_text().google_translate(source_lang=args.source, to_lang=args.to)
    elif args.auto == False:
      Processor().set_path(ppt_path).extract_text()
  elif task == "merge":
    Processor().set_path(ppt_path).replace_text()

if __name__ == "__main__":
  main()