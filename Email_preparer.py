import os
from Email_send import SendMail

# ------------------------------------------------------ settings ------------------------------------------------
root_dir = os.path.abspath(os.path.dirname(__file__))

emails = {
     '511': ['511e86@ibkconstructiongroup.com', 'sg@ibkconstructiongroup.com'],
     '262': ['dk@ibkconstructiongroup.com', 'vb@ibkconstructiongroup.com', '262kent@ibkconstructiongroup.com'],
     '28': ['tm@ibkconstructiongroup.com', '161e28@ibkconstructiongroup.com'],
     '199': ['199mineola@ibkconstructiongroup.com'],
     'manager': ['patricks@ibkconstructiongroup.com', 'ds@ibkconstructiongroup.com', 'yd@ibkconstructiongroup.com'],
     'office': ['timurp@ibkconstructiongroup.com', 'lilianas@ibkconstructiongroup.com']}

cc = ['cesarr@ibkconstructiongroup.com'] + emails['manager'] + emails['office']
# ------------------------------------------------------------------------------------------------------------------


def send_email(location):

    files_dir = os.path.join(root_dir, location)
    path_file = os.listdir(files_dir)

    list_pdf = []
    for names in path_file:
        if names[-4:] == '.pdf':
            pdf_path = os.path.join(files_dir, names)
            list_pdf.append(pdf_path)
    to = emails[location]
    active = SendMail(to=to, cc=cc, list_pdf=list_pdf)
    active.run()


if __name__ == "__main__":
    print('Only to be use with SafetyLogMain.py ')
    input('')

    # test_location = '262 199 28 511'.split()
    # for location in test_location:
    #     send_email(location=location)
