import os

from openpyxl import load_workbook

def get_user_info(xlsx_file):
    workbook = load_workbook(filename='data.xlsx')
    sheet = workbook.active
    user_info = []
    for row in sheet.rows:
        user = {}
        user['name'] = row[4].value
        if user['name'] is None or not user['name'].strip():
            continue
        user['surname'] = row[3].value
        email = row[6].value
        if email is None or not email.strip():
            continue
        user['email'] = email
        user['username'] = email[0: email.find('@')]
        user['workplace'] = row[9].value
        password = row[13].value
        if password is None or not str(password).strip():
            continue
        user['password'] = password

        user_info.append(user)
    return user_info


def create_rdp_file(filename, destdir, text):
    root_dirname = os.path.join(os.getcwd(), 'RDP CLOUD')
    if not os.path.exists(root_dirname):
        os.mkdir(root_dirname)
    dest_dirname = os.path.join(root_dirname, destdir)
    if not os.path.exists(dest_dirname):
        os.mkdir(dest_dirname)
    with open(os.path.join(dest_dirname, filename), 'w',  encoding='utf8') as file_writer:
        file_writer.write(text)


with open('template.txt', 'r', encoding='utf8') as file_reader:
    template_text = file_reader.read()

user_info = get_user_info(xlsx_file='data.xlsx')
for user in user_info:
    print(user)
    text = template_text.replace('%PATTERN%', user['username'])
    rdp_filename = user['surname']+user['name']+'.rdp'
    create_rdp_file(rdp_filename, user['workplace'], text)