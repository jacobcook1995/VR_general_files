import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

version = {}
with open("VR_general_files/version.py") as fp:
    exec(fp.read(), version)

setuptools.setup(
    name="VR_general_files",
    version=version["__version__"],
    author="Jacob Cook",
    author_email="j.cook17@imperial.ac.uk",
    description="Program to map the files of the Virtual Rainforest project.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/jacobcook1995/VR_general_files",
    packages=["VR_general_files"],
    entry_points={
        "console_scripts": [
            "get_the_files=VR_general_files.scripts:_VR_file_search_cli"
        ]
    },
    license="MIT",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
)
