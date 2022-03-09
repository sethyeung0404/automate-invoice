import os
import jinja2
import pandas as pd
import pdfkit

path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
options = {"page-width": "174", "page-height": "247"}


class GenerateInvoice:
    def __init__(self):
        pass

    def create_jinja2_template_html(self):
        data = pd.read_excel("data.xlsx")
        rows = data.shape[0]
        data = data.fillna("")

        templateLoader = jinja2.FileSystemLoader(searchpath="./")
        templateEnv = jinja2.Environment(loader=templateLoader)
        template_name = "template.html"
        template = templateEnv.get_template(template_name)

        current_dir = os.getcwd()
        output_dir = os.path.join(current_dir, "output/")
        if not os.path.exists(output_dir):
            os.mkdir(output_dir)
        os.chdir(output_dir)

        for i in range(rows):
            Client = data["Client"][i]
            Candidate = data["Candidate"][i]
            CandidateFullName = data["Candidate Full Name"][i]
            Position = data["Position / Content"][i]
            InvoiceNumber = data["Invoice #"][i]
            Address1 = data["Company Address 1"][i]
            Address2 = data["Company Address 2"][i]
            Address3 = data["District"][i]
            InvoiceDate = data["Invoice Date"][i].strftime("%d %b %Y")
            PaymentTerms = data["Terms"][i]
            DueDate = data["Due Date"][i].strftime("%d %b %Y")
            CommencementDate = data["Commencement Date"][i].strftime("%d %b %Y")
            Terms1 = data["Terms1"][i]
            Terms2 = data["Terms2"][i]
            Terms3 = data["Terms3"][i]
            Terms4 = data["Terms4"][i]
            InvoiceAmount = "{:.2f}".format(data["Actual Billing (HKD)"][i])

            html = template.render(
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
            fo = open(
                "OCG - " + str(Client) + " - ( " + str(InvoiceNumber) + " ).html",
                "w",
                encoding="UTF-8",
            )
            fo.writelines(html)
            fo.close()
            os.chdir(output_dir)
            print(
                f"create OCG - {str(Client)} - ( {str(InvoiceNumber)} ).html:",
                "完成",
            )
            infile = "OCG - " + str(Client) + " - ( " + str(InvoiceNumber) + " ).html"
            outfile = "OCG - " + str(Client) + " - ( " + str(InvoiceNumber) + " ).pdf"
            pdfkit.from_file(infile, outfile, configuration=config, options=options)
            print(
                f"create OCG - {str(Client)} - ( {str(InvoiceNumber)} ).pdf:",
                "完成",
            )


if __name__ == "__main__":
    GenerateInvoice().create_jinja2_template_html()
