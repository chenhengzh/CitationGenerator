# CitationGenerator

> **CitationGenerator**: The tool is designed for Zlin's students to simplify citation management. 


## Features 🎯

CitationGenerator has two modes, controlled by setting `GetPDF` in `docx_gen.py`. You can set `GetPDF=True` to crawl citation PDF files and generate the Word document (note that some PDF files require database access and cannot be crawled). Alternatively, you can set `GetPDF=False` to generate the Word document directly using only local citation PDF files.

I advise you to **run the `docx_gen.py` with `GetPDF=True` first** to download part of the PDF files and get an initial Word document. Then, **download the remaining PDFs manually using the web hyperlinks in the Word document**. After that, **run with `GetPDF=False`** to generate the final Word document. This way, you won't even need to manually adjust the hyperlinks in the Word document.

## Installation Guide 📥

It is recommended to create a Python virtual environment with [conda](https://conda.io/projects/conda/en/latest/user-guide/install/index.html) to install CitationSummary.

```shell
git clone https://github.com/chenhengzh/CitationSummary.git
cd CitationSummary
conda create -n citsum python=3.10
conda activate citsum
pip install -r requirements.txt
```

## Quick Start 🚀
- Visit serpapi to apply for an API key. Note that free version users are subject to a monthly access limit.
- Configure the `config.py` file according to your needs.
- To crawl citation information for a specific paper, run `paper_crawler.py`.
- To obtain all citation information for a particular author, run `author_crawler.py`.
