import re
from unidecode import unidecode


def normalize_name(name):
    return name.replace(" ", "").lower()

def normalize_phrase(name):
    new_name = unidecode(name)
    new_name = re.sub(r'\W+', '_', new_name)
    new_name = re.sub(r'__', '_', new_name)
    new_name = new_name.strip('_') + ".pdf"
    return new_name