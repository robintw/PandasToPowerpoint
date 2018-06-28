from setuptools import setup, find_packages
import os
import sys

CURR_DIR = os.path.abspath(os.path.dirname(__file__))

INSTALL_REQUIRES = [
    'pandas',
    'six',
    'python-pptx']

exec(open('pd2ppt/_version.py').read())

setup(
    name='pd2ppt',
    version=__version__,
    description='Python utility to take a Pandas DataFrame and create a '
                'Powerpoint table',
    url='https://github.com/robintw/PandasToPowerpoint',
    license='BSD-3-Clause',
    packages=find_packages(),
    install_requires=INSTALL_REQUIRES,
    zip_safe=False)
