START_DATE = '2018-01-18'
END_DATE = '2018-01-25'

CURATOR_EMAIL = 'your_email'
CURATOR_PASS = 'email_pass'
CURATOR_NAME = 'Your name'

# True\False
SEND_TO_PARENTS = True

MESSAGE_SUBJECT = "Фоксфорд. Отчёт об успеваемости"
MESSAGE_TEMPLATE = 'Здравствуйте!\nПрикрепляю отчёт за неделю\n' \
                   'С уважением, {}'.format(CURATOR_NAME)

CHILDREN = [
    {
        'name': 'Dima',
        'parent_email': 'mama_dimy@no.ru',
        'user_id': '1144089'
    },
    {
        'name': '',
        'parent_email': '',
        'user_id': ''
    },
    {
        'name': '',
        'parent_email': '',
        'user_id': ''
    }
]

