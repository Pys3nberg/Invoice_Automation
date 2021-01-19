from tempfile import NamedTemporaryFile
import shutil
import openpyxl
import csv


# Define some globals

# Should be moved to its own directory
INVOICE_WORKBOOK = "Invoices 2019-2020.xlsx"
PENDING_INVOICES = "pending invoices.csv"
DATE = "F11"
INVOICE_NO = "F12"
DESCRIPTION = "B18"
AMOUNT = "F18"
HOURLY_RATE = 8


def get_next_invoice(workbook):

    # Get the next invoice number by looping through the sheet names, and extracting the number
    # using list comprehension we make a list of these numbers and the find the max and add 1
    invoice_no = max(
        [int(inv.split(" ")[1]) for inv in workbook.sheetnames
         if not inv == "Template"]) + 1

    return invoice_no


if __name__ == "__main__":

    # open the work book that will have new invoices
    wb = openpyxl.load_workbook(INVOICE_WORKBOOK)
    # get the template sheet
    template = wb["Template"]
    # also create a temp file to write out all pending invoice entries, the temp file is now open
    tempFile = NamedTemporaryFile('w', delete=False, newline='')

    # now we open the pending invoices file as fid. also notice that tempfile is added at the end
    # this just means that it will be closed (with tempFile) does the clean up for us
    with open(PENDING_INVOICES, "r") as fid, tempFile:

        # init the two csv files, one for reading and then the temp file to write too
        pending_csv = csv.reader(fid, delimiter=',')
        temp_csv = csv.writer(tempFile, delimiter=',')

        # get headers and write to the temp csv
        headers = list(next(pending_csv))
        temp_csv.writerow(headers)

        for row in pending_csv:
            # loop through all rows and un-bundle the variables
            date, amount, customer_id, invoiceNo = row
            # if there is nothing in the invoice number column, it needs to be processed
            if not invoiceNo:

                # Get the next invoice number
                invoiceNo = get_next_invoice(wb)

                # Now we make a new sheet based off of the template sheet
                newInvoiceSheet = wb.copy_worksheet(template)
                # set the title of the sheet to the new invoice
                newInvoiceSheet.title = f"Invoice {invoiceNo:03}"
                # now we put the data in the right places
                newInvoiceSheet[DATE] = date
                newInvoiceSheet[INVOICE_NO] = f"{invoiceNo:03}"
                newInvoiceSheet[DESCRIPTION] = f"Care service - {int(amount) / HOURLY_RATE} hours - week {date}"
                newInvoiceSheet[AMOUNT] = float(amount)

            # write the row from pending invoices to the temp file with the updated invoice number column
            temp_csv.writerow([date, amount, customer_id, invoiceNo])

    # Save the invoice workbook with the new sheets
    wb.save(INVOICE_WORKBOOK)
    # Now replace the pending invoices csv with the temp file. This will have the updated invoice numbers now
    shutil.move(tempFile.name, PENDING_INVOICES)
