# CitationGenerator

> **CitationGenerator**: This tool is designed for Zlin's students to simplify citation management. 

## Features 🎯

CitationGenerator has two modes, controlled by setting `GetPDF` in `docx_gen.py`. You can set `GetPDF=True` to crawl citation PDF files and generate a Word document for each paper (note that some PDF files require database access and cannot be crawled). Alternatively, you can set `GetPDF=False` to generate Word documents directly using only local citation PDF files.

I advise you to **run `docx_gen.py` with `GetPDF=True` first** to download part of the PDF files and get an initial Word document. Then, **download the remaining PDFs manually from the web hyperlinks in Word documents**. After that, **run with `GetPDF=False`** to generate the final Word document. This way, you won't even need to manually adjust the hyperlinks in Word documents.

## Installation Guide 📥

It is recommended to create a Python virtual environment with [conda](https://conda.io/projects/conda/en/latest/user-guide/install/index.html) to install CitationGenerator.

```shell
git clone https://github.com/chenhengzh/CitationGenerator.git
cd CitationGenerator
conda create -n citgen python=3.10
conda activate citgen
pip install -r requirements.txt
```

## Quick Start 🚀

- Put citation data in `paper_list/`, such as `paper_list/papername/data/`
- Set `GetPDF=True` in `docx_gen.py` and run `docx_gen.py`. This may take several hours, so using tmux is recommended.
- Download the remaining PDFs manually in `paper_list/papername/`
- Set `GetPDF=False` and run `docx_gen.py` again to update the hyperlinks in Word documents.
