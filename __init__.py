""" init """
__version__ = '0.1.9'
import gettext
from .main import parse
from .main import parse_document_d2p
from .main import word2quiz
gettext.install('base', 'locales')  # install _() globally
