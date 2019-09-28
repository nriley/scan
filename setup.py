"""
Usage:
    python setup.py py2app
"""

from setuptools import setup

APP = ['update_dates_in_place.py']
DATA_FILES = []
OPTIONS = dict(
    plist=dict(
        CFBundleIdentifier='net.sabi.UpdateDatesInPlace',
        CFBundleName='Update Dates in Place',
    ))

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
