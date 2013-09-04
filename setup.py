from setuptools import setup, find_packages
setup(
    name = "HEIGVD_TimetableParser",
    version = "0.1",
    packages = find_packages(),

    install_requires = ['icalendar>=3.5', 'xlrd>=0.9.2'],

    # metadata for upload to PyPI
    author = "Leeroy Brun",
    author_email = "leeroy.brun@gmail.com",
    description = "Transforme un horaire au format XLS provenant de l'intranet du département FEE de la HEIG-VD en un fichier ICS.",
    license = "MIT",
    keywords = "heig-vd ics xls fee",
    url = "https://github.com/leeroybrun/heigvd-timetable-parser", 
)