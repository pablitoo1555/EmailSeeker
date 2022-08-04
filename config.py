from datetime import datetime, timedelta

# Subject that we want to search inbox for
subject = 'Consolidated And Position Statement COB'

# Folder where emails are saved in Outlook
trg_folder = 'Swap'

#output_dir = 'H:\\DEPT\\TRUSTADMN\\Liability management\\Swap Daily PDFs\\2022'
output_dir = 'Output'

latest_date = datetime.today().date()