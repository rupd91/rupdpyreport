# rupdpyreport/
# ├── setup.py
# ├── rupdpyreport/
# │   └── __init__.py

# python3 -m pip install -e .

from setuptools import setup, find_packages

setup(
    name='rupdpyreport',
    version='0.1',
    description='RUPD PowerPoint Reporting Tools',
    author='Rahul Upadhyay',
    packages=find_packages(),
    install_requires=[
        'python-pptx>=0.6.21',
        'pptxtopdf>=0.0.2',
        'Pillow>=8.0.0'
    ],
)
