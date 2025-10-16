#!/usr/bin/env python3
import optparse
import sys
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

options = None

def parse_args():
  parser = optparse.OptionParser(
    """
    'docxtofa' is a tool used to translate docx files to txt files with formatting for Fur Affinity.
    This can be helpful for posting stories in the description without having to lose formatting.
    Note: This software cannot recognize tables or images.

    run with: '{Python Path} docxtofa.py [file]' or just drag your docx file(s) onto the program
    
    Tool created by: Haika and Lexario
    """,
    version="%prog 1.2")

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
    textFilename = fileNoExtension + " (BBC).txt"
  else:
    print("Error: Passed in file is not docx format.")
    quit()

  return textFilename


def has_horizontal_line(paragraph):
  for run in paragraph._element.findall('.//w:r', namespaces=paragraph._element.nsmap):
    pict = run.find('.//w:pict', namespaces=paragraph._element.nsmap)
    if pict is not None:
      vrect = pict.find('.//v:rect', namespaces=paragraph._element.nsmap)
      if vrect is not None and vrect.get('{urn:schemas-microsoft-com:office:office}hr') == 't':
        return True
  return False


def extract_hyperlinks(paragraph):
  hyperlinks = []
  for child in paragraph._element:
    if child.tag == qn('w:hyperlink'):
      r_id = child.get(qn('r:id'))
      if r_id is None:
        continue

      # Resolve the URL
      part = paragraph.part
      if r_id not in part.rels:
        continue
      url = part.rels[r_id]._target

      # Extract display text
      text_fragments = []
      for node in child.iter():
        if node.tag == qn('w:t'):
          text_fragments.append(node.text)
      display_text = ''.join(text_fragments)

      hyperlinks.append((display_text, url))
  return hyperlinks


def check_lines_and_links(outFile, paragraph):
  if has_horizontal_line(paragraph):
    outFile.write('-----')
      
  links = extract_hyperlinks(paragraph)
  if links:
    for text, url in links:
      outFile.write(f"[url={url}]{text}[/url]")


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
    text = "[color=" + color + "]" + text + "[/color]"
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
    check_lines_and_links(outFile, paragraph)
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
