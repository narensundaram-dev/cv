import os
import re
import glob
import shutil
import zipfile
import logging
import argparse
from io import BytesIO
from datetime import datetime as dt

import nltk
import pandas as pd
import win32com.client
from xml.etree.ElementTree import XML

from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import TextConverter
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter


__author__ = "Narendran G"
__maintainer__ = "Narendran G"
__contact__ = "narensundaram007@gmail.com / +91 8678910063"

log = logging.getLogger(__file__.split('/')[-1])

# Loading all the available indian names
indian_names = open("names.txt", "r").read().lower()
indian_names = set(indian_names.split())


def config_logger(args):
    """
    This method is used to configure the logging format.

    :param args: script argument as `ArgumentParser instance`.
    :return: None
    """
    log_level = logging.INFO if args.log_level and args.log_level == 'INFO' else logging.DEBUG
    log.setLevel(log_level)
    log_handler = logging.StreamHandler()
    log_formatter = logging.Formatter('%(levelname)s: %(asctime)s - %(name)s:%(lineno)d - %(message)s')
    log_handler.setFormatter(log_formatter)
    log.addHandler(log_handler)


class CVReader(object):

    cwd = os.getcwd()

    __doc, __docx, __pdf = "doc", "docx", "pdf"
    docs_supported = (__doc, __docx, __pdf)

    def __init__(self, path):
        self.cwd = CVReader.cwd
        self.initial_path = path
        self.path = path
        self.text = ""

    @property
    def initial_filename(self):
        return os.path.split(self.initial_path)[-1]

    @property
    def filename(self):
        return os.path.split(self.path)[-1]

    @property
    def extension(self):
        return self.filename.split(".")[-1]

    def tokenize(self):
        try:
            text = self.text.encode('ascii', 'ignore').decode("ascii", "ignore")
            lines = [el.strip() for el in text.split("\n") if len(el) > 0]
            lines = [nltk.word_tokenize(el) for el in lines]
            lines = [nltk.pos_tag(el) for el in lines]
            sentences = nltk.sent_tokenize(text)
            sentences = [nltk.word_tokenize(sent) for sent in sentences]
            tokens = sentences
            sentences = [nltk.pos_tag(sent) for sent in sentences]
            dummy = []
            for el in tokens:
                dummy += el
            tokens = dummy
            return tokens, lines, sentences
        except Exception as e:
            log.error("Error on tokenizing the ")

    def extract_name(self):
        name = ""
        other_name_hits = []
        name_hits = []
        try:
            tokens, lines, sentences = self.tokenize()
            grammar = r'NAME: {<NN.*><NN.*><NN.*>*}'
            parser = nltk.RegexpParser(grammar)
            all_chunked_tokens = []
            for tagged_tokens in lines:
                if len(tagged_tokens) == 0:
                    continue
                chunked_tokens = parser.parse(tagged_tokens)
                all_chunked_tokens.append(chunked_tokens)
                for subtree in chunked_tokens.subtrees():
                    if subtree.label() == 'NAME':
                        for ind, leaf in enumerate(subtree.leaves()):
                            if leaf[0].lower() in indian_names and 'NN' in leaf[1]:
                                hit = " ".join([el[0] for el in subtree.leaves()[ind:ind + 3]])
                                if re.compile(r'[\d,:]').search(hit):
                                    continue
                                name_hits.append(hit)
                if len(name_hits) > 0:
                    name_hits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in name_hits]
                    name = " ".join([el[0].upper() + el[1:].lower() for el in name_hits[0].split() if len(el) > 0])
                    other_name_hits = name_hits[1:]

        except Exception as e:
            log.error("Error getting the name from the document.")
            return "", []

        name = " ".join(name.split(" ")[:2]) if len(name.split(" ")) >= 2 else name
        other_name_hits = other_name_hits[:5] if len(other_name_hits) >= 5 else other_name_hits
        return name, other_name_hits

    def extract_mobile(self):
        mobiles = set()
        pattern_mobile = [
            r"(\+91-)?(\+91)?(-\s)?([0-9]{3}).?([0-9]{3}).?([0-9]{4})",
            r"(\+91-)?(\+91)?(-\s)?([0-9]{4}).?([0-9]{3}).?([0-9]{3})"
        ]
        for pattern in pattern_mobile:
            matches = re.finditer(pattern, self.text, re.MULTILINE)
            for match in matches:
                mobiles.add(match.group())
        return ", ".join(list(mobiles))

    def extract_email(self):
        emails = set()
        pattern_email = r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)"
        matches = re.finditer(pattern_email, self.text, re.MULTILINE)
        for match in matches:
            emails.add(match.group().lower())
        emails = list(filter(lambda x: x.split(".")[-1] in ("com", "co", "in", "org"), list(emails)))
        return ", ".join(emails)

    def extract(self, text):
        self.text = text

        txt_file_path = self.filename.replace(self.extension, "txt")
        path = os.path.join(CVManager.path_txt_files, txt_file_path)
        with open(path, "w+") as f:
            f.write(self.text.encode("ascii", "ignore").decode("ascii", "ignore"))

        name, others = self.extract_name()
        return {
            "file_name": self.initial_filename,
            "name": name,
            "mobile": self.extract_mobile(),
            "email": self.extract_email(),
            "name_hints": ", ".join(others)
        }

    def doc2docx(self):
        filename = os.path.split(self.path)[-1]
        if "~$" not in filename:
            destination = os.path.join(CVManager.path_doc2docx_files, filename.replace(".doc", ".docx"))
            word = win32com.client.Dispatch("Word.application")
            document = word.Documents.Open(self.path)
            try:
                document.SaveAs2(destination, FileFormat=16)
                log.debug("Doc: {} converted to Docx: {}".format(self.path, destination))
                return destination
            except BaseException as e:
                log.error('Failed to Convert: {}\n.Error: {}'.format(self.path, e))
            finally:
                document.Close()
        return ""

    def read_doc(self):
        self.path = self.doc2docx()
        if self.path:
            return self.read_docx()
        else:
            return {}

    def read_docx(self):
        log.info("Reading: {}".format(self.initial_path))
        namespace = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        para = namespace + 'p'
        text = namespace + 't'

        document = zipfile.ZipFile(self.path)
        paragraphs = []
        for segment in ("word/header1.xml", "word/header2.xml", "word/header3.xml", "word/document.xml"):
            if segment in list(document.NameToInfo.keys()):
                xml = document.read(segment)
                tree = XML(xml)

                for paragraph in tree.iter(para):
                    texts = [n.text for n in paragraph.iter(text) if n.text]
                    if texts:
                        paragraphs.append(''.join(texts))
        document.close()

        text = '\n'.join(paragraphs)
        return self.extract(text)

    def read_pdf(self):
        log.info("Reading: {}".format(self.initial_path))
        manager = PDFResourceManager()
        layout = LAParams(all_texts=True)

        with BytesIO() as io:
            with TextConverter(manager, io, laparams=layout) as device:
                with open(self.path, "rb") as file_:
                    interpreter = PDFPageInterpreter(manager, device)
                    text = ""
                    for page in PDFPage.get_pages(file_, check_extractable=True):
                        interpreter.process_page(page)
                        text += io.getvalue().decode("ascii", "ignore")
        return self.extract(text)

    def read(self):
        try:
            if self.extension == CVReader.__doc:
                return self.read_doc()
            elif self.extension == CVReader.__docx:
                return self.read_docx()
            elif self.extension == CVReader.__pdf:
                return self.read_pdf()
            else:
                log.info("Skipping file: {}, The file is not in any of the formats({}).".format(
                    self.path, ", ".join(CVReader.docs_supported)))
                return {}
        except Exception as err:
            log.error("Error reading the file: {}. Please contact developer to fix it.".format(self.filename))


class CVManager(object):

    path_txt_files = os.path.join(os.getcwd(), "txts")
    path_doc2docx_files = os.path.join(os.getcwd(), "doc2docx")

    def __init__(self, args):
        self.args = args
        self.data = []

    @classmethod
    def setup(cls):
        os.makedirs(cls.path_txt_files, exist_ok=True)
        os.makedirs(cls.path_doc2docx_files, exist_ok=True)

    @classmethod
    def cleanup(cls):
        shutil.rmtree(CVManager.path_txt_files, ignore_errors=True)
        shutil.rmtree(CVManager.path_doc2docx_files, ignore_errors=True)

    def get(self):
        path = os.path.join(self.args.destination, "*")
        paths_all = glob.glob(path, recursive=True)
        paths = list(filter(lambda x: "~$" not in x, paths_all))

        for path in paths:
            data = CVReader(path).read()  # returns {"file_name": "", "name": "", "mobile": "", "email": ""}
            if data:
                self.data.append(data)
        return self

    def save(self):
        df = pd.DataFrame(self.data)
        path = os.path.join(os.getcwd(), "cv_info.xlsx")
        df.to_excel(path, index=False)


def get_args():
    arg_parser = argparse.ArgumentParser()
    arg_parser.add_argument('-f', '--destination', type=str,
                            help='Enter the folder path where you want to save the file.')
    arg_parser.add_argument('-log-level', '--log_level', type=str, choices=("INFO", "DEBUG"),
                            default="INFO", help='Where do you want to post the info?')
    return arg_parser.parse_args()


def main():
    args = get_args()
    config_logger(args)

    start = dt.now().strftime("%d-%m-%Y %H:%M:%S %p")
    log.info("Script starts at: {}".format(start))

    CVManager.setup()
    manager = CVManager(args).get()
    manager.save()
    CVManager.cleanup()

    end = dt.now().strftime("%d-%m-%Y %H:%M:%S %p")
    log.info("Script ends at: {}".format(end))


if __name__ == "__main__":
    main()
