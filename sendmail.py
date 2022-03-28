from email.message import EmailMessage


import mimetypes
import xlrd

loc = ("./mail.xlsx")

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
row_count = sheet.nrows
column_count = sheet.ncols

mime_type, _ = mimetypes.guess_type('2022-ozgur-kayabasi-cv.pdf')
mime_type, mime_subtype = mime_type.split('/')



for x in range(row_count):
    message = EmailMessage()
    mail = sheet.cell_value(x, 1)
    print(mail)
    sender = "ozgur.kayabasiii@gmail.com"
    recipient = mail
    message['From'] = sender
    message['To'] = recipient

    message['Subject'] = '2022 Zorunlu Yaz Staj'

    body = """
    Merhaba, 

    Ben Özgür Kayabaşı, Karadeniz teknik üniversitesi yazılım mühendisliği son sınıf öğrencisiyim. Zorunlu olarak yaz aylarında 60 günlük iş yeri eğitimi yapabileceğim ve ilgi alanlarımda kendimi geliştirebileceğim firmalar aramaktayım. CV dosyam ektedir. İnceleyip geri dönüşünüzü rica ederim. 

    Teşekkürler,
    İyi çalışmalar
    """
    message.set_content(body)

    with open('2022-ozgur-kayabasi-cv.pdf', 'rb') as file:
        message.add_attachment(file.read(),
        maintype=mime_type,
        subtype=mime_subtype,
        filename='2022-ozgur-kayabasi-cv.pdf')
    

    import smtplib
    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')

    mail_server.login("ozgur.kayabasiii@gmail.com", 'xxxx')

    mail_server.set_debuglevel(1)
    mail_server.send_message(message)
    mail_server.quit()
    del message['To']




