"""
Get a document, from source directory, with all paycheck of every workers and split it in many PDF one for every worker, the pdf can has one or more pages

"""
### LIBRARY
import smtplib
import fitz
import sys
import time
from os import path, makedirs, rename

### VARIABLES AND CONSTATNS
# Dictionary for translate month to number 
month_to_int = {
  "gennaio": "01",
  "febbraio": "02",
  "marzo": "03",
  "aprile": "04",
  "maggio": "05",
  "giugno": "06",
  "luglio": "07",
  "agosto": "08",
  "settembre": "09",
  "ottobre": "10",
  "novembre": "11",
  "dicembre": "12"
  }

# Email variables
smtp_server = "mx.yourmx.com"
smtp_port = 25
email_from = "from@yourmail.com"
email_to = "to@yourmail.com"
email_suject = "Subject email"
email_messag = ""

### FUNCTION
def occour_error(error_message):
  send_email(error_message)
  sys.exit(error_message)

def send_email(email_message):
  s = smtplib.SMTP(smtp_server, smtp_port)
  message = 'Subject: {}\n\n{}'.format(email_suject, email_message)
  s.sendmail(email_from, email_to, message)
  s.quit()

def get_words(words, row, start_point, end_point):
  string = ''
  
  # Iterate all words and return the word specified by row, start and end points 
  for i in words:
    # Get first and last names
    if(i[3] == row and (i[0] >= start_point and i[0] < end_point)):
      string += i[4].replace("'","")
  
  return string


# Create or merge a pdf with a desired page numeber from the original doc
def save_file(file, doc, page_number):
  # If file exsists than it append current page
  if(path.exists(file)):
    with fitz.open(file) as output_doc: # Open a file
      output_doc.insert_pdf(doc, from_page=page_number, to_page=page_number) # Append current page of main document
      output_doc.save(file, incremental=True, encryption=0) # Save file
  # otherwise create a new file with current page
  else:
    with fitz.open() as output_doc: # Open an empty file
      output_doc.insert_pdf(doc, from_page=page_number, to_page=page_number) # Append current page of main document
      output_doc.save(file) # Save file


def make_document(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3):

  # Check if the source file exists and open or abort programm
  if(not path.exists(source_filename)):
    occour_error("[ERROR] - File is not found !")
  
  with fitz.open(source_filename) as doc: # Open file

    # Iterate all pages in PDF
    for page in doc.pages(0, doc.page_count): # Range of pages doc.pages(START, STOP, STEP) | doc.page_count 
      words_of_page = page.get_text("words")  # Create a list of items that contain the words of the current page

      # If the page is empty, skip it
      if(len(words_of_page) <= 1):
        continue

      # Iterate all words of the current page to get first and last names, month and year 
      for i in words_of_page:
        name = get_words(words_of_page, name_row, name_p1, name_p2).lower() # Name of worker
        month = month_to_int[get_words(words_of_page, data_row, data_p1, data_p2).lower()] # Month of the year
        year = get_words(words_of_page, data_row, data_p2, data_p3) # Year of the paycheck
      
      # Combine the name, month and year to create the destination folder and output filename
      if(not __debug__): # If used the -O options the __debug__ is False
        destination_folder = '/home/m0rp30/tmp/'
      else:
        destination_folder = '/usr/share/pydio/data/personal/' + name + '/'
      output_filename = destination_folder + name + '_' + month + '_' + year + '.pdf'
      
      # If destination folder doesn't exists make it
      if(not path.exists(destination_folder)):
        makedirs(destination_folder)
      save_file(output_filename, doc, page.number) # Save or append current page in the pdf

  if(not __debug__): # If used the -O options the __debug__ is False
    backup_filename = "/home/m0rp30/backup/" + path.basename(source_filename).split('.')[0] + "_" + year + "_" + month + ".pdf"
  else:
    backup_filename = "/mnt/cedolini/backup/" + path.basename(source_filename).split('.')[0] + "_" + year + "_" + month + ".pdf"

  try:
    rename(source_filename, backup_filename)
  except Exception as e:
    occour_error(e)


### MAIN
if __name__ == "__main__":

  # CEDOLINI
  source_filename = sys.argv[1] #'/mnt/cedolini/CEDOLINI_SOCI.pdf'
  name_row = 163.51426696777344 # line_no 11
  name_p1 = 63.300010681152344
  name_p2 = 316.49993896484375
  data_row = 122.79430389404297 # line no 7  # That is the bottom position of row's
  data_p1 = 341.81988525390625
  data_p2 = 405.1197509765625
  data_p3 = 449.4296569824219
  make_document(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3)
  
  # STACED
  source_filename = sys.argv[2] #'/mnt/cedolini/STACED_SOCI.pdf'
  name_row = 63.390953063964844 # line_no 3
  name_p1 = 272.998046875
  name_p2 = 582.5955200195312
  data_row = 51.390953063964844 # line_no 2
  data_p1 = 20.99901580810547
  data_p2 = 323.3979797363281
  data_p3 = 582.5955200195312
  make_document(source_filename, name_row, name_p1, name_p2, data_row, data_p1, data_p2, data_p3)
  
  send_email("Elaborazione CEDOLINI/STACED completata senza errori!")
