from setuptools import setup, find_packages
setup(
    name = "HEIGVD_TimetableParser",
    version = "0.1",
    packages = find_packages(),

    install_requires = ['icalendar>=3.5', 'xlrd>=0.9.2'],

    # metadata for upload to PyPI
    author = "Leeroy Brun",
    author_email = "leeroy.brun@gmail.com",
    #description = "Tool used to remove all archives stored inside an Amazon Glacier vault.",
   # license = "MIT",
    #keywords = "aws amazon glacier boto archives vaults",
    #url = "https://github.com/leeroybrun/glacier-vault-remove", 
)