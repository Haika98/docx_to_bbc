#!/usr/bin/env python3
import optparse
import sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

options = None

def parse_args():
  parser = optparse.OptionParser(
    """
    'docxtofa' is a tool used to translate docx files to txt files with formatting for Fur Affinity.
    This can be helpful for posting stories in the description without having to lose formatting.
    Note: This software cannot recognize links or horizontal lines. You'll have to place in 5+ dashes (-) manually. 

    run with: '{Python Path} docxtofa.py [file]' or just drag your docx file(s) onto the program
    
    Tool created by: Haika and Lexario
    """,
    version="%prog 1.0")

  # Available command line arguments
  parser.add_option("-v", "--verbose", dest="verbose", action="store_true", default=False, help="Enable this flag for additional print logging.")
  parser.add_option("-q", "--quiet", dest="quiet", action="store_true", default=False, help="Enable this flag to turn off any print logging.")
  
  global options
  (options, arguments) = parser.parse_args()

  # Throw error if user does not specify required parameters
  if options.verbose and options.quiet:
      parser.error("Do not set verbose and quiet flags simultaneously.")
  return options


def get_outFilename(inFile):
  fileExtension = inFile.split(".")[-1]
  fileNoExtension = inFile[:-(len(fileExtension)+1)]
  if fileExtension == "docx":
    textFilename = fileNoExtension + " (FA).txt"
  else:
    print("Error: Passed in file is not docx format.")
    quit()

  return textFilename


def parse_run(outFile, run):
  text = run.text
  if run.underline:
    text = "[u]" + text + "[/u]"
  if run.italic:
    text = "[i]" + text + "[/i]"
  if run.bold:
    text = "[b]" + text + "[/b]"
  if run.font.strike:
    text = "[s]" + text + "[/s]"
  if run.font.superscript:
    text = "[sup]" + text + "[/sup]"
  if run.font.subscript:
    text = "[sub]" + text + "[/sub]"
  if run.font.color.type != None:
    color = "#" + str(run.font.color.rgb)
    text = "[color=" + color + "]" + text + "[/" + color + "]"
  outFile.write(text)


def parse_paragraph(outFile, paragraph):
  for run in paragraph.runs:
    parse_run(outFile, run)


def parse_alignment(outFile, paragraph):
  if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER: # Center aligned
    outFile.write("[center]")
    parse_paragraph(outFile, paragraph)
    outFile.write("[/center]")
  elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT: # Right aligned
    outFile.write("[right]")
    parse_paragraph(outFile, paragraph)
    outFile.write("[/right]")
  else: # Left (default) aligned
    parse_paragraph(outFile, paragraph)
  outFile.write("\n")


def parse_doc(inFile, outFile):
  doc = docx.Document(inFile)
  for paragraph in doc.paragraphs:
    parse_alignment(outFile, paragraph)


########## MAIN ##########
def main():
  options = parse_args() # Parse optional arguments
  inFiles = []
  for arg in sys.argv[1:]: # Store input files
    if arg[0] != '-':
      inFiles.append(arg)

  if not options.quiet:
    print("Ready to translate these files:")
    for p in inFiles:
      print(p)
    print("Text files will appear in the same directory as the source.")
    x = input("Press Enter to continue or q to quit...")
    if x == 'q' or x == 'Q':
      quit()
    print("")

  for p in inFiles:
    if options.verbose:
      print("Parsing path:", p)
    outFilename = get_outFilename(p)
    if options.verbose:
      print("Creating text file with name:", outFilename.split("\\")[-1])
    outFile = open(outFilename, "w", encoding="utf-8")
    if options.verbose:
      print("Parsing document:", p.split("\\")[-1])
    parse_doc(p, outFile)
    outFile.close()
    if options.verbose:
      print(outFilename, "Written successfully\n")

  if not options.quiet:
    print("File(s) translated.")
    input("Press Enter to continue...")
########## END MAIN ##########

# TODO:
# Horizontal lines
# Headings
# Hyperlinks (URLs)
# Detect if last line is new line

if __name__ == '__main__':
  main()
