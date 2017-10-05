from setuptools import setup, find_packages
from codecs import open
from os import path

here = path.abspath(path.dirname(__file__))

# Get the long description from the README file
with open(path.join(here, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name='sample',

    version='1.0.0',

    description='Figure out the distance between two zipcodes and dump the result in a spreadsheet',
    long_description=long_description,

    # The project's main homepage.
    url='https://github.com/esologic/zipcode-distance-excel',

    # Author details
    author='Devon Bray',
    author_email='dev@esologic.com',

    # Choose your license
    license='MIT',

    classifiers=[
        'Development Status :: 3 - Alpha',

        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
    ],

    # What does your project relate to?
    keywords='zipcodes postcodes excel',

    packages=find_packages(exclude=['contrib', 'docs', 'tests']),

    py_modules=["zip-distance-xl"],

    install_requires=['pyzipcode3', 'geopy', 'openpyxl'],
)