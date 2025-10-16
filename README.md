# docx-to-bbc
A tool used to translate a docx file to a txt file for sites using BBCode formatting. This can be helpful for posting stories in the description without having to lose formatting.

Run with: `python3 docxtobbc.py [options] "FILE.docx"`

Output file will be: `FILE (BBC).txt`

-----

Notes:

* Requires python3 and python-docx (Use `pip install python-docx`)
* Current tool limitations:
  * Headings will output as normal text.
  * Bulleted lines will output as unbulleted lines.
  * Tables and images unfortunately cannot be parsed and a blank line will be left in its place.

Website Specific: 

* *Fur Affinity* and *Inkbunny* do not support indenting. This program will output them fine to the text file, but they won't mean anything when actually used on those sites.
* *Fur Affinity* and *Inkbunny* do not support bulleted lists. Each bulleted item will go on its own line, but it will not contain the bullets, dashes, or indentation.
* Inkbunny does not support superscripts or subscripts. Tags such as [sup] and [sub] will be added, but will not 

-----

Tool created by: Haika and Lexario
