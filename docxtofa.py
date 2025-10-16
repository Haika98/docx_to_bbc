#!/usr/bin/env python3
import optparse
import sys
import io
import os
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH

options = None

def parse_args():
  parser = optparse.OptionParser(
    """
    'docxtofa' is a tool used to translate a docx file to a txt with formatting for Fur Affinity.
    This can be helpful for posting stories in the description without having to lose formatting.
    Note: This software cannot recognize horizontal lines. You'll have to place in 5+ dashes (-) manually. 

    run with: 'python3 docxtofa.py [file]' or just drag your docx file(s) onto the program
    
    Tool created by: Haika and Lexario
    """,
    version="%prog 0.5")

  # Available command line arguments
  parser.add_option("-v", "--verbose", action="store_true", default=False, dest="verbose", help="Enable this flag for additional print logging.")
  parser.add_option("-q", "--quiet", action="store_true", default=False, dest="quiet", help="Enable this flag to turn off any print logging.")
  
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
    textFilename = fileNoExtension + ".txt"
  return textFilename


def parse_doc(inFile, outFile):
  doc = docx.Document(inFile)
  for paragraph in doc.paragraphs:
    outFile.write("\t")
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
      outFile.write("[center]")
      for run in paragraph.runs:
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
        outFile.write(text)
      outFile.write("[/center]")
      outFile.write("\n")
    elif paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
      outFile.write("[right]")
      for run in paragraph.runs:
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
        outFile.write(text)
      outFile.write("[/right]")
      outFile.write("\n")
    else:
      for run in paragraph.runs:
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
        outFile.write(text)
      outFile.write("\n")


########## MAIN ##########
options = parse_args() # Parse optional arguments
inFiles = []
for arg in sys.argv[1:]: # Store input files
  if arg[0] != '-':
    inFiles.append(arg)

for p in inFiles:
  if options.verbose:
    print("Parsing path:", p)
  outFilename = get_outFilename(p)
  if options.verbose:
    print("Creating text file with name:", outFilename.split("/")[-1])
  outFile = open(outFilename, "w", encoding="utf-8")
  parse_doc(p, outFile)
  outFile.close()

if not options.quiet:
  print("Successfully translated files.")
  os.system("pause")
########## END MAIN ##########

# TODO:
# Horizontal lines
# Headings
# Colors
# Hyperlinks (URLs)
# Detect if last line is new line
