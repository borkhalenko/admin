from openpyxl import load_workbook

def get_user_info(xlsx_file):
    workbook = load_workbook(filename='data.xlsx')
    sheet = workbook.active
    
    user_info = []

    for row in sheet.rows:
        user = {}
        user['name'] = row[0].value
        user['surname'] = row[1].value
        email = row[2].value
        user['email']=email
        user['username'] = email[0: email.find('@')]
        user['workplace']=row[3].value
        user_info.append(user)
    return user_info

with open('template.txt', 'r') as file_reader:
    template_text = file_reader.read()

print(type(template_text))
user_info = get_user_info(xlsx_file='data.xlsx')
for user in user_info:
    text = template_text.replace('%PATTERN%', user['username'])
    rdp_filename = user['surname']+user['name']+'.rdp'
    with open(rdp_filename, 'w') as new_file:
        new_file.write(text)
