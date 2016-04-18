import os
import sys

import daogao

try:
    from setuptools import setup, Extension
except ImportError:
    from distutils.core import setup, Extension
    pass

if sys.argv[-1] == 'publish':
    os.system('python setup.py sdist upload')
    sys.exit()

packages = [
    'daogao'
]

requires = ['tornado', 'psycopg2', 'openpyxl',
			'requests', 'pillow', ]
setup(
    name='daogao',
    version=daogao.__version__,
    description='Excel generation daemon.',
    long_description=open('README.txt').read(),
    author='Gamaliel Espinoza Macedo',
    author_email='gamaliel.espinoza@gmail.com',
    url='https://bitbucket.org/gamikun/daogao',
    packages=packages,
    package_dir={'daogao': 'daogao'},
    install_requires=requires,
    include_package_data=True,
    package_data={},
    ext_modules=[],
    zip_safe=False,
    classifiers=(
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Natural Language :: English',
        'Programming Language :: Python :: 2.7',
    ),
)