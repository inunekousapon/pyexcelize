from __future__ import annotations

from glob import glob
from os.path import basename
from os.path import splitext

from setuptools import Extension
from setuptools import setup
from setuptools import find_packages


def _requires_from_file(filename):
    return open(filename).read().splitlines()


setup(
    name="pyexcelize",
    version="0.2.1",
    license="MIT",
    description="A Python library for reading and writing Excel files",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    author="Junichi Yukawa",
    author_email="j.yukawa@gmail.com",
    ext_modules=[Extension('excelize', ['main.go'])],
    build_golang={'root': 'github.com/asottile/fake', 'strip': False},
    setup_requires=['setuptools-golang'],
    packages=find_packages('./'),
    install_requires=_requires_from_file('requirements.txt'),
    zip_safe=False,
    include_package_data=True,
    py_modules=[splitext(basename(path))[0] for path in glob('./*.py')],
)