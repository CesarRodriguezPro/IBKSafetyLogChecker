
from email import encoders
from email.mime.multipart import MIMEMultipart, MIMEBase
from email.mime.text import MIMEText
import datetime, smtplib, os


class SendMail:

    def __init__(self, to, cc, list_pdf):
        self.user_name = 'cesarr'
        self.user_password = 'SuperConcrete01'
        self.today = datetime.datetime.today().strftime('%m/%d/%Y')
        self.emailUser = "cesarr@ibkconstructiongroup.com"
        self.to = to
        self.cc = cc
        self.list_pdf = list_pdf

    def run(self):
        with open('message.txt', 'r') as message_file:
            message = message_file.read()

        msg = MIMEMultipart()
        msg['From'] = self.emailUser
        msg['To'] = ', '.join(map(str, self.to))
        msg['Cc'] = ', '.join(map(str, self.cc))
        msg['Subject'] = "Automatic Report for TimeStation"

        body = message
        msg.attach(MIMEText(body, 'plain'))

        if self.list_pdf is not None:
            for each_file_path in self.list_pdf:
                try:
                    file_name = os.path.basename(each_file_path)
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(open(each_file_path, "rb").read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', 'attachment', filename=file_name)
                    msg.attach(part)
                except:
                    print("could not attache file")

        server = smtplib.SMTP('smtp.ibkconstructiongroup.com')
        server.starttls()
        server.login(self.user_name, password=self.user_password)
        text = msg.as_string()
        server.sendmail(self.emailUser, self.to + self.cc, text)
        server.quit()
        print('message was send')


if __name__ == '__main__':

    print('This is only to be used from the EmployeesMissing.py')
    input('')