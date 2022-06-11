""" main module """
import gettext


def get_translator():

    en = gettext.translation('base', localedir='locales', languages=['en'])
    nl = gettext.translation('base', localedir='locales', languages=['nl'])

    nl.install()
    nl.gettext("message to translate")  # no warning
    print(_("message to translate"))    # no warning

    # return nl.gettext  # uncommenting this line changes warning line 24 to 'Unexpected argument'

    language = ['nl']
    if 'nl' in language[0]:
        return nl.gettext
    return en.gettext


def a_method():
    _ = get_translator()
    print(_("message to translate"))
