"""Simple excel report drawing framework based on pandas

The main idea here is to have a plotter-like object (called a drawer),
that moves through an excel sheet matrix and draws objects that are supplied to it.
"""

# Always prefer setuptools over distutils
from setuptools import setup, find_packages
# To use consistent encoding
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.md'), encoding = 'utf-8') as f:
    long_description = f.read()
    
setup(
    name = 'xlsxpandas',

    # Versions should comply with PEP440. For a discussion on single-sourcing
    # the version across setup.py and the project code, see
    # https://packaging.python.org/en/latest/single_source_version.html

    version = '0.0.1',
    
    description = 'Simple excel report drawing framework based on pandas',
    long_description = long_description,
    
    # The project's main homepage
    url = 'https://github.com/sztal/xlsxpandas',
    
    # Author details
    author = 'Szymon Talaga',
    author_email = 'stalaga@protonmail.com',
    
    # Choose your licence
    licence = 'MIT',
    
    # see https://pypi.python.org/pypi?%3Aaction=list_classifiers
    classifiers = [
        # How mature is this project?
        #   3 - Alpha
        #   4 - Beta
        #   5 - Production/Stable
        'Development Status :: 3 - Alpha',
        
        # Intended audience
        'Intended Audience :: Developers :: Data Scientist :: Bussiness Analysts',
        'Topic :: Reporting',
        
        # License
        'License :: OSI Approved :: MIT License',
        
        # Specify the Python versions you support here. In particular, ensure
        # that you indicate whether you support Python 2, Python3 or both.
        'Programming Language :: Python :: 3.6',
    ],
    
    # What does your project relate to?
    keywords = 'excel reporting automation',
    
    # You can just specify the packages manually here if your project is simple.
    packages = find_packages(exclude = ['contrib', 'docs', 'tests'])
)
