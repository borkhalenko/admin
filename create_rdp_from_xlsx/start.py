import os

from openpyxl import load_workbook

def get_rdp_template(template_filename, username, appname, pseudonym):
    with open('template_rdp.txt', 'r', encoding='utf8') as file_reader:
        rdp_file_template = file_reader.read()
    rdp_file_template = rdp_file_template.replace('%USERNAME%', username)
    rdp_file_template = rdp_file_template.replace('%APPNAME%', appname)
    rdp_file_template = rdp_file_template.replace('%PSEUDONYM%', pseudonym)
    return rdp_file_template

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
        password = row[16].value
        if password is None or not str(password).strip():
            continue
        user['password'] = password
        send_by_email = row[17].value
        if not send_by_email is None and (str(send_by_email).upper() == 'YES' or str(send_by_email).upper() == 'ДА'):
            user['send_by_email'] = True
        else:
            user['send_by_email'] = False
        medoc_access = row[18].value
        if not medoc_access is None and (str(medoc_access).upper() == 'YES' or str(medoc_access).upper() == 'ДА'):
            user['MEDOC_ACCESS'] = True
        else:
            user['MEDOC_ACCESS'] = False
        user_info.append(user)
    return user_info


def create_rdp_file(filename, destdir, text):


    os.makedirs(destdir, exist_ok=True)
    full_path = os.path.join(destdir, filename)
    with open(full_path, 'w',  encoding='utf8') as file_writer:
        file_writer.write(text)
    return full_path


def send_email(template_file, attachments, user_info):

    import smtplib, ssl
    from email import encoders
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    with open(template_file, 'r', encoding='utf8') as file_reader:
        template_text = file_reader.read()

    template_text = template_text.replace('%USERNAME%', f"{user_info['name']} {user_info['surname']}")
    template_text = template_text.replace('%PASSWORD%', user_info['password'])

    message = MIMEMultipart()
    message["From"] = 'admin@metal.kiev.ua'
    message['To'] = user_info['email']
    message['Subject'] = 'Доступ до хмарного серверу Metal Holding Trade'
    message.attach(MIMEText(template_text, "plain"))

    for att_file in attachments:
        basefilename = os.path.basename(att_file)
        with open(att_file, "rb") as attachment:
            part = MIMEApplication(attachment.read(), Name = basefilename)
        part['Content-Disposition'] = f'attachment; filename={basefilename}'
        message.attach(part)
    
    text = message.as_string()
    context = ssl.create_default_context()
    context.check_hostname = False
    context.verify_mode = ssl.CERT_NONE
    with smtplib.SMTP_SSL('mail.metal.kiev.ua', 465, context=context) as server:
        server.login('admin@metal.kiev.ua', '3voImDyYnGA$')
        server.sendmail('admin@metal.kiev.ua', 'o.borkhalenko@metal.kiev.ua', text)




users_info = get_user_info(xlsx_file='data.xlsx')

for user in users_info:
    files = []
    print(user)
    
    rdp_filename = f"1C_{user['surname']}_{user['name']}.rdp"
    appname = '1cestart'
    pseudonym = '1C Предприятие'
    text = get_rdp_template('template_rdp.txt', user['username'], appname, pseudonym)
    rdp_path = os.path.join(os.getcwd(), 'RDP_CLOUD', user['workplace'])
    rdp_full_filepath = create_rdp_file(rdp_filename, rdp_path, text)
    files.append(rdp_full_filepath)



    if user['MEDOC_ACCESS']:
        rdp_filename = f"MEDOC_{user['surname']}_{user['name']}.rdp"
        appname = 'Station'
        pseudonym = 'M.E.Doc Station'
        text = get_rdp_template('template_rdp.txt', user['username'], appname, pseudonym)
        rdp_path = os.path.join(os.getcwd(), 'RDP_CLOUD', user['workplace'])
        rdp_full_filepath = create_rdp_file(rdp_filename, rdp_path, text)
        files.append(rdp_full_filepath)
    if user['send_by_email']:
        print("SENDING EMAIL.......")
        if user['surname'] == 'Борхаленко':
            send_email('template_mail.txt', files, user)