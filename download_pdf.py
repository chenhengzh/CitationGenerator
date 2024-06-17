import logging
import requests
import arxiv
from fuzzywuzzy import fuzz

Timeout=30

class Paper:
    def __init__(self, title='', info='', abstract='', PDF='', filename='', link='') -> None:
        self.title = title
        self.info = info
        self.abstract = abstract
        self.PDF = PDF
        self.filename = filename
        self.link = link

    def display(self):
        logging.info("+++===================================================================================================+++")
        logging.info(f"title: {self.title}")
        logging.info(f"filename: {self.filename}")
        logging.info(f"info: {self.info}")  
        logging.info(f"abstract: {self.abstract}")
        if self.PDF == '':
            logging.info("no PDF resource.")
        else:
            logging.info(f"PDF: {self.PDF}")
        logging.info(f"paper_link: {self.link}")

def are_strings_almost_matching(string1, string2, threshold=90):
    # 使用 fuzz.ratio() 比较字符串相似性
    similarity_ratio = fuzz.ratio(string1, string2)
    return similarity_ratio >= threshold

def download_pdf_in_arxiv(title, abstract, save_path):
    # 使用 arxiv API 搜索文章
    client = arxiv.Client()
    search = arxiv.Search(query=f'ti:{title} AND abs:{abstract}')
    try:
        result = next(client.results(search))

    except StopIteration:
        logging.info("find nothing.")
        return False

    # 检查是否有搜索结果
    if result:
        ismatch = are_strings_almost_matching(title, result.title, 85)
        if ismatch:
            try:
                result.download_pdf(dirpath="", filename=save_path)
                logging.info(f"PDF has been downloaded: {save_path}")
                return True
            except Exception as e:
                logging.info(f"Network error. {e}")
                return False
        else:
            logging.info("find nothing.")
            return False
    else:
        logging.info("find nothing.")
        return False


def download_pdf_if_exists(url, save_path):
    try:
        # 发送 HTTP 请求，添加超时时间
        response = requests.get(url, stream=True, timeout=Timeout)

        # 检查响应状态码和内容类型
        if response.status_code == 200 and 'application/pdf' in response.headers.get('Content-Type', ''):
            # 如果是PDF文件，保存到本地
            with open(save_path, 'wb') as pdf_file:
                for chunk in response.iter_content(chunk_size=128):
                    pdf_file.write(chunk)
            logging.info(f"PDF has been downloaded: {save_path}")
            return True
        else:
            logging.info("Download failed, maybe the paper does not provide a PDF file.")
            return False

    except requests.Timeout:
        logging.info("Request timed out.")
        return False

    except requests.RequestException as e:
        logging.info(f"Network error: {e}")
        return False


def get_pdf(paper, pth):
    link = paper.PDF
    if not paper.PDF:
        link = paper.link
    if link:
        logging.info("+==============try to get pdf from link=============+")
        flag = download_pdf_if_exists(link, pth)
        if not flag:
            logging.info("+==============search pdf in arxiv=============+")
            flag = download_pdf_in_arxiv(paper.title, paper.abstract, pth)
    else:
        logging.info("+==============search pdf in arxiv=============+")
        flag = download_pdf_in_arxiv(paper.title, paper.abstract, pth)

    return flag