from setuptools import setup

APP = ['CampusContacts.py']
DATA_FILES = ['./supporting_files/contacts.xlsx']
OPTIONS = {
 'iconfile':'cru.icns',
 'argv_emulation': True,
 'packages': ['certifi','selenium', 'openpyxl'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
