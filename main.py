from Processor import Processor
from Processor import Paths

def main():
  
  demo_path = r"C:\Users\YourUsername\Paste\your\pptx\file\path\here"

  # The following takes your pptx file and converts all the text into an excel
  Processor().use_path(demo_path).extract_text()
  
  # Open the excel
  # Original text will be on column A
  # Paste/write all your translations on column B
  # Once done, you can merge these changes back into the pptx file 
  # It will be saved as a new file in the same folder

  # Uncomment the below and comment the extract_text() line
  # Processor().use_path(demo_path).replace_text()


if __name__ == "__main__":
  main()