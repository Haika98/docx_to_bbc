#!/usr/bin/env python3
import optparse
import sys
import docx
import os

options = None

def parse_args():
  parser = optparse.OptionParser(
    """
    This software is provided \"AS IS\", without warranty of any kind.\n
    'docxtofa' is a tool used to translate a docx file to a txt with formatting for Fur Affinity.
    This can be helpful for posting stories in the description without having to lose formatting.

    run with: 'python3 docxtofa.py [file]' or just drag your docx file onto the program
    Tool created by: Haika
    """,
    version="%prog 1.0")

  #Available command line arguments
  parser.add_option("-v", "--verbose", action="store_true", default=False, dest="verbose", help="Enable this flag for additional print logging.")
  parser.add_option("-q", "--quiet", action="store_true", default=False, dest="quiet", help="Enable this flag to turn off any print logging.")
  
  global options
  (options, arguments) = parser.parse_args()

  #Throw error if user does not specify required parameters
  if options.verbose and options.quiet:
      parser.error("Do not set verbose and quiet flags simultaneously.")
  return options

########## MAIN ##########
options = parse_args() #Option parser
inFiles = sys.argv[1:]
for p in inFiles:
  print(p)

os.system("pause")
########## END MAIN ##########
