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
        "lxml>=4.9",
        "latex2mathml>=3.77.0",
        "pywin32>=306",
        "pypdfium2>=5.0",
        "Pillow>=10.0",
    ],
    entry_points={
        "console_scripts": [
            "md2word=md2word.cli.main:main",
        ],
    },
    python_requires=">=3.8",
)
