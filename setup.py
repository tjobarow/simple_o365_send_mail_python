from setuptools import setup

setup(
    name="simple_o365_send_mail",
    version="1.00",
    py_modules=["simple_o365_send_mail"],
    install_requires=[
    ],
    author="Thomas Obarowski",
    author_email="tjobarow@gmail.com",
    description="A wrapper making it easier to send emails via O365/Msgraph API",
    long_description=open('README.md','r').read(),
    long_description_content_type="text/markdown",
    url="",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.10",
)
