from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="filewitch",
    version="0.2.1",
    author="Raktim Kalita",
    author_email="raktimkalita.ai@gmail.com",
    description="A Python library for converting files between different formats",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Rktim/filewitch",
    packages=find_packages(),
    license="MIT",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.7",
    install_requires=[
        "pandas>=1.5.0",
        "openpyxl>=3.1.0",
        "python-docx>=0.8.11",
        "click>=8.0.0",
        "python-pptx>=0.6.21",
        "reportlab>=4.0.4",
        "Pillow>=10.0.0",
        "docx2pdf>=0.1.8"
    ],
    entry_points={
        "console_scripts": [
            "filewitch=filewitch.cli:main",
        ],
    },
)