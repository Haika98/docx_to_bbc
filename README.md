# docx-to-fa
A tool used to translate a docx file to a txt with formatting for Fur Affinity. This can be helpful for posting stories in the description without having to lose formatting.

Run with: `python3 docxtofa.py [options] "FILE.docx"`

-----

Notes:

* Requires python3 and python-docx (Use `pip install python-docx`)
* Fur Affinity does not support indenting. This program will output them fine to the text file, but they won't mean anything when actually used on the site.
* Current tool limitations:
  * Headings will output as normal text.
  * Bulleted lines will output as unbulleted lines.
  * Hyperlinks unfortunately cannot be parsed and a blank line will be left in its place.
  * Horizontal lines cannot be parsed and a blank line will be left in its place. You'll have to insert five or more dashes (-----) manually where you'd like them to go.

-----

Tool created by: Haika and Lexario
