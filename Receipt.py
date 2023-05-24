import os
import jinja2
import pandas as pd
import pdfkit

path_wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
options = {"page-width": "174", "page-height": "247"}


class GenerateReceipt:
    def __init__(self):
        pass

    def create_jinja2_template_html(self):
        data = pd.read_excel("data.xlsx")
        rows = data.shape[0]
        data = data.fillna("")

        templateLoader = jinja2.FileSystemLoader(searchpath="./")
        templateEnv = jinja2.Environment(loader=templateLoader)
        template_name = "Receipt-template.html"
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
            Position = data["Position"][i]
            InvoiceNumber = data["Invoice #"][i]
            ReceiptNumber = data["Receipt #"][i]
            PaymentMethod = data["Payment Method"][i]
            PaymentReceivedDate = data["Payment Received Date"][i].strftime("%d %b %Y")
            InvoiceAmount = "{0:,.2f}".format(data["Amount Received from Invoice(HKD)"][i])
            Address1 = data["Company Address 1"][i]
            Address2 = data["Company Address 2"][i]
            Address3 = data["District"][i]

            html = template.render(
                Candidate=Candidate,
                Client=Client,
                CandidateFullName=CandidateFullName,
                InvoiceNumber=InvoiceNumber,
                Position=Position,
                Address1=Address1,
                Address2=Address2,
                Address3=Address3,
                InvoiceAmount=InvoiceAmount,
                PaymentMethod=PaymentMethod,
                PaymentReceivedDate=PaymentReceivedDate,
                ReceiptNumber=ReceiptNumber,
                
            )
            fo = open(
                "OCG - " + str(Client) + " - ( " + str(ReceiptNumber) + " ).html",
                "w",
                encoding="UTF-8",
            )
            fo.writelines(html)
            fo.close()
            os.chdir(output_dir)
            print(
                f"create OCG - {str(Client)} - ( {str(ReceiptNumber)} ).html:",
                "完成",
            )
            infile = "OCG - " + str(Client) + " - ( " + str(ReceiptNumber) + " ).html"
            outfile = "OCG - " + str(Client) + " - ( " + str(ReceiptNumber) + " ).pdf"
            pdfkit.from_file(infile, outfile, configuration=config, options=options)
            print(
                f"create OCG - {str(Client)} - ( {str(ReceiptNumber)} ).pdf:",
                "完成",
            )


if __name__ == "__main__":
    GenerateReceipt().create_jinja2_template_html()
