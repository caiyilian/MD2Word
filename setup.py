from setuptools import setup, find_packages

setup(
    name="md2word",
    version="0.1.0",
    description="Convert Markdown files to Word (.docx) documents",
    packages=find_packages(),
    install_requires=[
        "mistune>=3.0",
        "python-docx>=1.1.0",
        "pyyaml>=6.0",
    ],
    entry_points={
        "console_scripts": [
            "md2word=md2word.cli.main:main",
            "docx2md=md2word.cli.main:docx2md_main",
        ],
    },
    python_requires=">=3.8",
)
