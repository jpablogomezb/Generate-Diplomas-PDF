import xlrd
import smtplib
import os
import sys
import datetime
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw

# Meant to convert full names to abbreviations
def shorten( text, _max ):
    t = text.split(" ")
    text = ''
    if len(t)>1:
        for i in t[:-1]:
            text += i[0] + '.'
    text += ' ' + t[-1]
    if len(text) < _max :
        return text
    else :
        return -1
        #return text

# Add name, institute and project to certificate
def make_certi( ID, name, workshop, w_x, w_y, date, signer):
    img = Image.open("template_diploma.jpg")
    draw = ImageDraw.Draw(img)
    # Load font
    font_name = ImageFont.truetype("Fonts/Lato-Regular.ttf", 100)
    font_workshop = ImageFont.truetype("Fonts/Lato-Regular.ttf", 65)
    font_date = ImageFont.truetype("Fonts/Lato-Light.ttf", 22)
    font_signed = ImageFont.truetype("Fonts/Lato-Light.ttf", 24)

    # Check sizes and if it is possible to abbreviate
    # if not the IDs are added to an error list
    if ( len( name ) > 25 ):
        name = shorten( name, 25 )
    if ( len( workshop ) > 100 ):
        workshop = shorten( workshop, 100 )
    if ( len( signer ) > 100 ):
        signer = shorten(signer, 100 )

    if name == -1 or workshop == -1 or signer == -1 :
        return -1
    else:
        # Insert text into image template
        # TEXT POSITIONS: Update here depending on the template (template_diploma.jpg)
        draw.text((335, 400), name, (0,0,0), font=font_name)
        # Update w_x and w_x inside the xls file (list-attendees.xls).
        draw.text((float(w_x), float(w_y)), workshop, (0,0,0), font=font_workshop )
        draw.text((585, 805), date, (0,0,0), font=font_date )
        draw.text((655  , 1050), signer, (0,0,0), font=font_signed )

        if not os.path.exists('PDFs') :
            os.makedirs('PDFs')

        # Save as a PDF
        try:
            img.save('PDFs/'+str(ID)+'.pdf', "PDF", resolution=100.0)
            return 'PDFs/'+str(ID)+'.pdf'
        except Exception as e:
            print(e)

# Email the certificate as an attachment
def email_certi( filename, receiver, workshop):
    username = "jpablogomezb.development" #EMAIL ACCOUNT (USERNAME) here
    password = "" #PASSWORD HERE
    sender = username + '@gmail.com'

    msg = MIMEMultipart() 
    msg['Subject'] = "Certificate: {}".format(workshop)
    msg['From'] = username+'@gmail.com'
    msg['Reply-to'] = username + '@gmail.com'
    msg['To'] = receiver

    # That is what u see if dont have an email reader:
    msg.preamble = 'Multipart massage.\n'

    # Body
    part = MIMEText( "Hello,\n\nPlease find attached your certificate.\n\nGreetings." )
    msg.attach(part)

    # Attachment
    part = MIMEApplication(open(filename,"rb").read())
    part.add_header('Content-Disposition', 'attachment', filename = os.path.basename(filename))
    msg.attach( part )

    # Login
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.starttls()
    server.login( username, password )

    # Send the email
    server.sendmail(msg['From'], msg['To'], msg.as_string())

if __name__ == "__main__":
    error_list = []
    error_count = 0

    os.chdir(os.path.dirname(os.path.abspath((sys.argv[0]))))

    # Read data from an excel sheet from row 2
    Book = xlrd.open_workbook('list-attendees.xls')
    WorkSheet = Book.sheet_by_name('Sheet1')

    num_row = WorkSheet.nrows - 1
    row = 0

    while row < num_row:
        row += 1

        ID = WorkSheet.cell_value( row, 0 )
        name = WorkSheet.cell_value( row, 1 )
        workshop = WorkSheet.cell_value( row, 2 )
        workshop_x = WorkSheet.cell_value(row, 3 )
        workshop_y = WorkSheet.cell_value(row, 4 )
        date = WorkSheet.cell_value( row, 5 )
        signer = WorkSheet.cell_value( row, 6 )
        receiver = WorkSheet.cell_value( row, 7)
        #date = datetime.strftime(fecha,'%b %d, %Y')
        # Make certificate and check if it was successful
        try:
            filename = make_certi( ID, name, workshop, workshop_x, workshop_y, date, signer)
            print(filename)
        except Exception as e:
            #print('error')
            print(e)

        # Successfully made certificate
        if filename != -1:
            try:
                email_certi( filename, receiver, workshop)
                print("Sent to, {}".format(receiver))
            except Exception as e:
                pass
                print(e)
        # Add to error list
        else:
            error_list.append( ID )
            error_count += 1

    #Print all failed IDs
    #print("%s Errors- List: , %s") %(error_count, error_list)
    #print("%s Errors") %(error_count)
