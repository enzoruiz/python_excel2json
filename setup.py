import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='python_excel2json',
    version='0.1',
    author="Enzo Ruiz Pelaez",
    author_email="enzo.rp.90@gmail.com",
    description="Python Excel to Json",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/enzoruiz/python_excel2json",
    packages=setuptools.find_packages(),
    install_requires=[
        'xlrd'
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
