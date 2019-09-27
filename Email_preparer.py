import os
from Email_send import SendMail


# ------------------------------------------------------ settings ------------------------------------------------
root_dir = os.path.abspath(os.path.dirname(__file__))

emails = {
     '511': ['511e86@ibkconstructiongroup.com', 'sg@ibkconstructiongroup.com'],
     '225': ['rr@ibkconstructiongroup.com','dk@ibkconstructiongroup.com', '225w28@ibkconstructiongroup.com'],
     '262': ['vb@ibkconstructiongroup.com', '262kent@ibkconstructiongroup.com'],
     '161': ['tm@ibkconstructiongroup.com', '161e28@ibkconstructiongroup.com'],
     '199': ['199mineola@ibkconstructiongroup.com'],
     '300': ['vb@ibkconstructiongroup.com'],
     '123': ['alexo@rcsrebar.com'],
     '215': ['225w28@ibkconstructiongroup.com'],
     '1230': ['vb@ibkconstructiongroup.com', '123madison@ibkconstructiongroup.com'],
     'manager': [ 'ds@ibkconstructiongroup.com', 'yd@ibkconstructiongroup.com','josephb@ibkconstructiongroup.com'],
     'office': ['timurp@ibkconstructiongroup.com', 'lilianas@ibkconstructiongroup.com'],
     }

# cc = ['cesarr@ibkconstructiongroup.com'] + emails['manager'] + emails['office']
cc = ['timurp@ibkconstructiongroup.com']

# ------------------------------------------------------------------------------------------------------------------


def send_email(location, total_employees):

    files_dir = os.path.join(root_dir, location)
    path_file = os.listdir(files_dir)

    list_pdf = []
    for names in path_file:
        if names[-4:] == '.pdf':
            pdf_path = os.path.join(files_dir, names)
            list_pdf.append(pdf_path)
    # to = emails[location]
    to = ['cesarr@ibkconstructiongroup.com']
    active = SendMail(to=to, cc=cc, list_pdf=list_pdf)
    active.run(total_employees=total_employees)


if __name__ == "__main__":
    print('Only to be use with SafetyLogMain.py ')
    input('')
