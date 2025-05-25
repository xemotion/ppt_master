# ppt_master
<ppt_master - create Tempalte>
Create a presentation template using the python-pptx library.
Analyze the slides and extract all text frames, then generate metadata using an LLM to define each element’s "role" and a "description" that explains the element’s structural purpose.
Reconstruct the template using the metadata’s fields: each key is mapped to the original text and replaced with the corresponding "role", maintaining position information to serve as a semantic identifier instead of a generic element ID.
Current limitation: 
* The program does not recognize text inside tables, grouped shapes, or the font color of shapes that contain nested text boxes.
Future improvements:
* Fix font color (default to black)
* Use only two predefined fonts
* Filter out shapes with very small font sizes (e.g., 10pt or smaller), which are likely used for annotations.
