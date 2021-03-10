from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
import os
import mimetypes


def template(template, file_name, p_name):
    """
    Takes template and set html variables values, setting E-mail body.
    """
    with open(f'{template}', 'r') as html:
        temp = Template(html.read())
        body = temp.substitute(trabalho=file_name, nome=p_name)
        return body


def get_path(file_name, path):
    """
    Gets file's path.
    """
    path_list = []
    if not file_name:
        return None
    for r, d, fs in os.walk(path):
        for f in fs:
            if file_name in f:
                f_nome, ext = os.path.splitext(f)
                cp_path = os.path.join(r, f)
                path_list.append((f_nome, ext, cp_path))
    return path_list


# Configure attachments.
def anexo(f_name, ext, cp_path):
    """
    Returns attachment MIME.
    """

    try:
        ctype, enconding = mimetypes.guess_type(f_name)

        if ctype is None or enconding is not None:
            ctype = 'application/octet-stream'

        maint, subt = ctype.split('/', 1)

        if maint == 'image':
            with open(f'{f_name}{ext}', 'rb') as img:
                mime = MIMEImage(img.read(), _subtype='jpg')

        else:
            with open(f'{cp_path}', 'rb') as f:
                mime = MIMEBase(maint, subt)
                mime.set_payload(f.read())
            encoders.encode_base64(mime)
        mime.add_header('Content-Disposition', 'attachment',
                        filename=f'{f_name}{ext}')
        return mime
    except Exception as e:
        print(f'Error: {e}')
        raise e


def set_template(template_name, file_name, p_name):
    """
    Insert template into the E-mail.
    """
    corp = template(template_name, file_name, p_name)
    msg = MIMEMultipart()
    msg['from'] = p_name
    msg['subject'] = f'{file_name}'
    msg.attach(MIMEText(corp, 'html'))
    return msg
