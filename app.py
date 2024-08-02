import os
import time
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import traceback
import xml.etree.ElementTree as ET
from datetime import datetime


def getFormattedDateTime():
    return datetime.now().strftime('%d.%m.%Y - %H:%M:%S')


def get_new_files(folder_path, from_date, first_check):
    if first_check or not hasattr(get_new_files, 'files_sent') or not get_new_files.files_sent:
        print(
            f"{getFormattedDateTime()} -> Перевіряю '{folder_path}' на наявність нових файлів")

    all_files = set(get_all_files(folder_path, from_date))
    prev_files = set(get_new_files.prev_files if hasattr(
        get_new_files, 'prev_files') else [])
    new_files = all_files - prev_files
    get_new_files.prev_files = list(all_files)

    if new_files:
        print(f"{getFormattedDateTime()} -> Нові файли знайдено:")
        for file in new_files:
            print(f"{file}")
        get_new_files.files_sent = False  # Reset the flag as new files are found
    else:
        get_new_files.files_sent = True  # No new files, set flag to true

    return new_files


def get_all_files(folder_path, from_date):
    all_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_mtime = os.path.getmtime(file_path)
            file_mtime_dt = datetime.fromtimestamp(file_mtime)
            if file_mtime_dt >= datetime.strptime(from_date, '%d.%m.%Y'):
                all_files.append(file_path)
    return all_files


def group_files_by_extension(files):
    file_dict = {}
    for file in files:
        _, ext = os.path.splitext(file)
        ext = ext.lower()[1:]  # Normalize extension to lowercase
        if ext not in file_dict:
            file_dict[ext] = []
        file_dict[ext].append(file)
    return file_dict


def monitor_folder_and_send_email(settings):
    print('---Системні повідомлення---')
    print(f"{getFormattedDateTime()} -> Програма успішно запущена. Натисніть Ctrl+C щоб зупинити програму")

    first_check = True  # Initial flag for first check

    while True:
        try:
            new_files = get_new_files(
                settings['monitor']['directory'], settings['monitor']['from_date'], first_check)
            first_check = False  # Reset the first check flag after the first iteration

            if new_files and settings['email']['send_messages']:
                file_dict = group_files_by_extension(new_files)
                send_email(file_dict, settings)
                first_check = True  # Reset the first check flag after sending files
        except Exception as exp:
            log_error(exp)
            print(
                f"{getFormattedDateTime()} -> Виникла помилка, перевірьте errors.txt ")
        time.sleep(settings['monitor']['interval'])


def send_email(file_dict, settings, batch_size=70, delay=5):
    try:
        send_info = get_send_info(file_dict, settings['recipients'])
        with smtplib.SMTP(settings['smtp']['server'], settings['smtp']['port']) as server:
            if settings['smtp']['use_auth'] == 'true':
                server.login(settings['smtp']['username'],
                             settings['smtp']['password'])
            for recipient_batch in batches(list(send_info.items()), batch_size):
                recipient_emails, recipient_infos = zip(*recipient_batch)
                if len(file_dict) > 0:
                    for ext, files in file_dict.items():
                        if ext != 'path':
                            for recipient_email, recipient_info in zip(recipient_emails, recipient_infos):
                                if all(file in recipient_info for file in files):
                                    msg = MIMEMultipart()
                                    msg['From'] = f"{settings['email']['sender_name']} <{settings['email']['sender_email']}>"
                                    msg['To'] = recipient_email
                                    msg['Subject'] = generate_subject(files)
                                    msg.attach(MIMEText(generate_email_body(
                                        settings['email']['message'], settings['email']['signature']), 'html'))

                                    for file in files:
                                        with open(file, "rb") as attachment:
                                            part = MIMEApplication(
                                                attachment.read(), Name=os.path.basename(file))
                                            part[
                                                'Content-Disposition'] = f'attachment; filename="{os.path.basename(file)}"'
                                            msg.attach(part)

                                    try:
                                        server.send_message(msg)
                                        message = f"{getFormattedDateTime()} -> Повідомлення відправлено до {recipient_email}, файли: {','.join([os.path.basename(f) for f in files])}"
                                        print(message)
                                        log_sent_messages(message)
                                        time.sleep(delay)
                                    except smtplib.SMTPException as e:
                                        log_error(e)
                                        print(
                                            f"{getFormattedDateTime()} -> Помилка відправки пошти до {recipient_email}. Перевірьте errors.txt для отримання інформації.")
    except Exception as expe:
        log_error(expe)
        print(f"{getFormattedDateTime()} -> Помилка відправки пошти. Перевірьте errors.txt для отримання інформації.")


def batches(iterable, n):
    iterable = list(iterable)  # Convert to list to support slicing
    for i in range(0, len(iterable), n):
        yield iterable[i:i + n]


def generate_subject(files):
    path_elements = files[0].split(os.path.sep)
    if len(path_elements) >= 5:
        return f"{path_elements[2]} : {path_elements[3]} ({path_elements[4]})"
    else:
        return "Файли RECON"


def generate_email_body(message, signature):
    html_body = f"""
    <html>
        <body>
        <h3>Доброго дня, колеги.</h3>
            <p>{message}</p>
            <hr>
            <blockquote><i>{signature}</i></blockquote>
        </body>
    </html>
    """
    return html_body


def get_send_info(file_dict, recipients):
    send_info = {}
    for ext, files in file_dict.items():
        for recipient in recipients:
            if recipient['email'] not in send_info:
                send_info[recipient['email']] = []
            for file in files:
                base_name = os.path.basename(file)
                file_name, file_extension = os.path.splitext(base_name)
                recon_number = str(extract_numeric_part(file_name))
                if recon_number in recipient['files'] or '*' in recipient['files']:
                    send_info[recipient['email']].append(file)
    return send_info


def extract_numeric_part(filename):
    return ''.join(c for c in filename if c.isdigit())


def log_sent_messages(message):
    with open('sent_messages.txt', 'a', encoding='UTF-8') as log_file:
        log_file.write(f"{message}\n")


def log_error(exception):
    with open('errors.txt', 'a', encoding='UTF-8') as error_file:
        error_file.write(f"Error: {exception}\n")
        error_file.write(f"Traceback: {traceback.format_exc()}\n")


def parse_settings():
    settings = {}
    try:
        tree = ET.parse('settings.xml')
        root = tree.getroot()

        smtp_settings = root.find('smtp')
        settings['smtp'] = {
            'server': smtp_settings.find('server').text,
            'port': int(smtp_settings.find('port').text),
            'use_auth': smtp_settings.find('use_auth').text.lower(),
            'username': smtp_settings.find('username').text,
            'password': smtp_settings.find('password').text,
            'use_SSL': smtp_settings.find('use_SSL').text.lower(),
        }

        email_settings = root.find('email')
        settings['email'] = {
            'sender_name': email_settings.find('sender_name').text,
            'sender_email': email_settings.find('sender_email').text,
            'signature': email_settings.find('signature').text,
            'message': email_settings.find('message').text,
            'send_messages': email_settings.find('send_messages').text.lower() == 'true',
        }

        recipients = root.find('recipients')
        settings['recipients'] = []
        for recipient in recipients.findall('recipient'):
            settings['recipients'].append({
                'email': recipient.get('email'),
                'files': recipient.get('files'),
            })

        monitor_settings = root.find('monitor')
        date_from_file = ''
        if monitor_settings.find('from_date').text != 'none':
            date_from_file = datetime.strptime(
                monitor_settings.find('from_date').text, '%d.%m.%Y')
        current_datetime = datetime.now()  # get current date
        from_date = ''
        if date_from_file:
            from_date = date_from_file.strftime('%d.%m.%Y')
        else:
            from_date = current_datetime.strftime('%d.%m.%Y')
        print('---Налаштування програми---')
        print(f'Файли будуть опрацьовані з: {from_date}')
        print("Директорія, що моніториться програмою: " +
              monitor_settings.find('directory').text+"\n")

        settings['monitor'] = {
            'directory': monitor_settings.find('directory').text,
            'interval': int(monitor_settings.find('interval').text),
            'from_date': from_date
        }
    except Exception as exp:
        print('Помилка при парсингу файлу налаштувань')
        log_error(exp)
    return settings


settings = parse_settings()
monitor_folder_and_send_email(settings)
