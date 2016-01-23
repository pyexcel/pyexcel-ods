try:
    from setuptools import setup, find_packages
except ImportError:
    from ez_setup import use_setuptools
    use_setuptools()
    from setuptools import setup, find_packages
import sys

with open("README.rst", 'r') as readme:
    README_txt = readme.read()

dependencies = [
    'pyexcel-io>=0.1.0',
    'odfpy==0.9.6',
]

extra_dependencies = {}

if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    dependencies.append('ordereddict')

setup(
    name='pyexcel-ods',
    author='C. W.',
    version='0.1.0',
    author_email='wangc_2011 (at) hotmail.com',
    url='https://github.com/pyexcel/pyexcel-ods',
    description='A wrapper library to read, manipulate and write data in ods format',
    install_requires=dependencies,
    extras_require=extras,
    packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
    include_package_data=True,
    long_description=README_txt,
    zip_safe=False,
    tests_require=['nose'],
    keywords=[
        'excel',
        'python',
        'pyexcel',
    ],
    license='New BSD',
    classifiers=[
        'Topic :: Office/Business',
        'Topic :: Utilities',
        'Topic :: Software Development :: Libraries',
        'Programming Language :: Python',
        'License :: OSI Approved :: BSD License',
        'Intended Audience :: Developers',
    ]
)