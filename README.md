# Generate-Diplomas-PDF
The script generates PDF diplomas from a list of attendees (spreadsheet), and they are sent by e-mail.

- Update with your jpg template in line 30 ("template_diploma.jpg")
- Insert attendess and workshop information in list-attendees.xls file. 
- Update text positions in relation to your diploma template in lines 52 - 56.
- Update also w_x and w_x (workshop name text position) inside the xls file (list-attendees.xls) depending on the workshop name string.
- To send diplomas by email: update username and password (from gmail account) in lines 70 and 71.

Requirements: xlrd; Pillow.
