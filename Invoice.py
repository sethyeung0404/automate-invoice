import os
import jinja2
import pandas as pd
import pdfkit

path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
options = {"page-width": "175mm", "page-height": "257mm", "zoom": "1.15"}


class GenerateInvoice:
    def __init__(self):
        pass

    def create_invoice(self):
        data = pd.read_excel("data.xlsx")
        rows = data.shape[0]
        data = data.fillna("")
        templateLoader = jinja2.FileSystemLoader(searchpath="./")
        templateEnv = jinja2.Environment(loader=templateLoader)

        template_invoice = "Invoice-template.html"
        template1 = templateEnv.get_template(template_invoice)
        template_invoice_noname = "Invoice-template_withoutName.html"
        template2 = templateEnv.get_template(template_invoice_noname)

        current_dir = os.getcwd()
        output_dir = os.path.join(current_dir, "output/")
        if not os.path.exists(output_dir):
            os.mkdir(output_dir)
        os.chdir(output_dir)

        for i in range(rows):
            Client = data["Client"][i]
            Candidate = data["Candidate"][i]
            CandidateFullName = data["Candidate Full Name"][i]
            Position = data["Position"][i]
            InvoiceNumber = data["Invoice #"][i]
            Address1 = data["Company Address 1"][i]
            Address2 = data["Company Address 2"][i]
            Address3 = data["District"][i]
            InvoiceDate = data["Invoice Date"][i].strftime("%d %b %Y")
            PaymentTerms = data["Terms"][i]
            DueDate = data["Due Date"][i].strftime("%d %b %Y")
            CommencementDate = data["Commencement Date"][i].strftime("%d %b %Y")
            CommencementDateFilename = data["Commencement Date"][i].strftime("%Y%m%d")
            Terms1 = data["Terms1"][i]
            Terms2 = data["Terms2"][i]
            Terms3 = data["Terms3"][i]
            Terms4 = data["Terms4"][i]
            InvoiceAmount = "{0:,.2f}".format(data["Amount Received from Invoice(HKD)"][i])

            Invoice = template1.render(
                Candidate=Candidate,
                Client=Client,
                CandidateFullName=CandidateFullName,
                InvoiceNumber=InvoiceNumber,
                Address1=Address1,
                Address2=Address2,
                Address3=Address3,
                InvoiceDate=InvoiceDate,
                PaymentTerms=PaymentTerms,
                DuteDate=DueDate,
                CommencementDate=CommencementDate,
                Position=Position,
                InvoiceAmount=InvoiceAmount,
                Terms1=Terms1,
                Terms2=Terms2,
                Terms3=Terms3,
                Terms4=Terms4,
            )
            Invoice2 = template2.render(
                Client=Client,
                InvoiceNumber=InvoiceNumber,
                Address1=Address1,
                Address2=Address2,
                Address3=Address3,
                InvoiceDate=InvoiceDate,
                PaymentTerms=PaymentTerms,
                DuteDate=DueDate,
                CommencementDate=CommencementDate,
                CommencementDateFilename=CommencementDateFilename,
                InvoiceAmount=InvoiceAmount,
            )

            fo = open(
                "OCG - "
                + str(InvoiceNumber)
                + " - "
                + str(Client)
                + " - "
                + str(Candidate)
                + ".html",
                "w",
                encoding="UTF-8",
            )
            fo.writelines(Invoice)
            fo.close()
            os.chdir(output_dir)
            print(
                f"create OCG - {str(InvoiceNumber)} - {str(Client)} - {str(Candidate)}.html:",
                "Done",
            )
            infile = (
                "OCG - "
                + str(InvoiceNumber)
                + " - "
                + str(Client)
                + " - "
                + str(Candidate)
                + ".html"
            )
            outfile = (
                "OCG - "
                + str(InvoiceNumber)
                + " - "
                + str(Client)
                + " - "
                + str(Candidate)
                + ".pdf"
            )
            pdfkit.from_file(infile, outfile, configuration=config, options=options)
            print(
                f"create OCG - {str(InvoiceNumber)} - {str(Client)} - {str(Candidate)}.pdf:",
                "Done",
            )
            fo = open(
                "OCG - "
                + str(InvoiceNumber)
                + " - "
                + str(Client)
                + " - "
                + str(CommencementDateFilename)
                + ".html",
                "w",
                encoding="UTF-8",
            )
            fo.writelines(Invoice2)
            fo.close()
            os.chdir(output_dir)
            print(
                f"create OCG - {str(InvoiceNumber)} - {str(Client)} - {str(CommencementDateFilename)}.html:",
                "Done",
            )
            infile = (
                "OCG - "
                + str(InvoiceNumber)
                + " - "
                + str(Client)
                + " - "
                + str(CommencementDateFilename)
                + ".html"
            )
            outfile = (
                "OCG - "
                + str(InvoiceNumber)
                + " - "
                + str(Client)
                + " - "
                + str(CommencementDateFilename)
                + ".pdf"
            )
            pdfkit.from_file(infile, outfile, configuration=config, options=options)
            print(
                f"create OCG - {str(InvoiceNumber)} - {str(Client)} - {str(CommencementDateFilename)}.pdf:",
                "Done",
            )

    
if __name__ == "__main__":
    GenerateInvoice().create_invoice()
