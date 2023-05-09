import extract_msg
import win32com.client
from docx import Document
from docx2pdf import convert

# Identify results email by "SITE NOT AFFECTED"
# Read E-mail body
# Extract: Client Ref
# Extract: Search Address from BackOffice and store as a variable to reference
# Extract: Search Type
# Create a document using "SF Headed Template" and insert Client Ref/Search Address/Search Type
# Insert: Template text - hyperlinking "Legal Notices" to bottom of document
# Convert to PDF
# Send to back office:
    # Subject line = SF****000
    # Sender = sfreturnedsearches@searchflow.co.uk
    # To: searchreturns@searchflowsearches.co.uk
