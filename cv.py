import os
import re
import glob
import json
import shutil
import zipfile
import logging
import argparse
import traceback
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

from utils.email_normalizer import EmailNormalizer


__author__ = "Narendran G"
__maintainer__ = "Narendran G"
__contact__ = "narensundaram007@gmail.com / +91 8678910063"

log = logging.getLogger(__file__.split('/')[-1])

# Loading all the available indian names
path_indian_names = os.path.join(os.getcwd(), "setup", "names.csv")
with open(path_indian_names, "r") as f:
    series = pd.read_csv(f)
    indian_names = series.drop_duplicates().values


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
        self.text_path = ""
        self.data = {}
        self.skip = False

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
            log.error("Error on tokenizing.")

    def extract_name(self):
        name = ""
        other_name_hits = []
        name_hits = []
        try:
            tokens, lines, sentences = self.tokenize()
            grammar = r'NAME: {<NN.*><NN.*><NN.*>*}'
            parser = nltk.RegexpParser(grammar)
            for tagged_tokens in lines:
                if len(tagged_tokens) == 0:
                    continue
                tokens_chunked = parser.parse(tagged_tokens)
                for subtree in tokens_chunked.subtrees():
                    # if subtree.label() == 'NAME':
                    for idx, leaf in enumerate(subtree.leaves()):
                        # if leaf[0].lower() in indian_names and 'NN' in leaf[1]:
                        # print("leaf: ", leaf)
                        if leaf[0].lower() in indian_names:
                            hit = " ".join([el[0] for el in subtree.leaves()[idx:idx + 3]])
                            if re.compile(r'[\d,:]').search(hit):
                                continue
                            name_hits.append(hit)
                if len(name_hits) > 0:
                    name_hits = [re.sub(r'[^a-zA-Z \-]', '', el).strip() for el in name_hits]
                    name = " ".join([el[0].upper() + el[1:].lower() for el in name_hits[0].split() if len(el) > 0])
                    other_name_hits = name_hits[1:]

        except BaseException as e:
            log.error("Error getting the name from the document.")
            return "", []

        names = [None] * 3
        name = name.replace(".", " ")
        name = name.split(" ")[:3] if len(name.split(" ")) >= 3 else name.split(" ")
        for i, n in enumerate(name):
            if len(n) <= 2 or n.lower() in indian_names:
                names[i] = n
        other_name_hits = other_name_hits[:5] if len(other_name_hits) >= 5 else other_name_hits
        return names, other_name_hits

    def extract_mobile(self):
        mobiles = set()
        pattern_mobile = [
            # r"(\+91-)?(\+91)?(-\s)?([0-9]{2,4}).?([0-9]{2,4}).?([0-9]{2,4})",
            r"(\+91-)?(\+91)?(-\s)?([0-9]{3}).?([0-9]{3}).?([0-9]{4})",
            r"(\+91-)?(\+91)?(-\s)?([0-9]{4}).?([0-9]{3}).?([0-9]{3})",
            r"(\+91-)?(\+91)?(-\s)?([0-9]{3}).?([0-9]{4}).?([0-9]{3})",
            r"(\+91-)?(\+91)?(-\s)?([0-9]{4}).?([0-9]{2}).?([0-9]{4})",
            r"(\+91-)?(\+91)?(-\s)?([0-9]{5}).?([0-9]{5})"
        ]
        for pattern in pattern_mobile:
            matches = re.finditer(pattern, self.text, re.MULTILINE)
            for match in matches:
                mobiles.add(match.group())
        return self.normalize_mobile_numbers(mobiles)

    def extract_email(self):
        email_ids = [None] * 2
        emails = set()
        pattern_email = r"([a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+)"
        matches = re.finditer(pattern_email, self.text, re.MULTILINE)
        for match in matches:
            email = EmailNormalizer(match.group().lower()).normalize()
            emails.add(email)
        emails = list(filter(lambda x: x.split(".")[-1] in ("com", "co", "in", "org"), list(emails)))
        emails = emails[:2] if len(emails) >= 2 else emails
        emails = list(set(emails))
        for idx, email in enumerate(emails):
            email_ids[idx] = email
        return email_ids

    def extract(self, text):
        self.text = text.strip()

        txt_file_path = self.filename.replace(self.extension, "txt")
        path = os.path.join(CVManager.path_txt_files, txt_file_path)
        with open(path, "w+") as f:
            self.text_path = path
            f.write(self.text.encode("ascii", "ignore").decode("ascii", "ignore"))

        names, others = self.extract_name()
        fname, mname, lname = names[0], names[1], names[2]
        mobile1, mobile2 = self.extract_mobile()
        email1, email2 = self.extract_email()
        return {
            "file_name": self.initial_filename,
            "first_name": fname,
            "middle_name": mname,
            "last_name": lname,
            "mobile1": mobile1,
            "mobile2": mobile2,
            "email1": email1,
            "email2": email2,
            # "name_hints": ", ".join(others)
        }

    @classmethod
    def normalize_mobile_numbers(cls, mobiles):
        mobile_numbers = [None] * 2
        mobiles = list(mobiles)[:2] if len(mobiles) >= 2 else list(mobiles)
        for idx, mobile in enumerate(mobiles):
            norm_mobile = mobile.replace("+91", "").replace(" ", "")
            norm_mobile = re.sub(r"\D", "", norm_mobile)
            if norm_mobile[0] in "06789":
                mobiles[idx] = int(norm_mobile)
        mobiles = list(map(lambda x: int(x), [n for n in map(lambda x: str(x), mobiles) if len(n) == 10]))
        mobiles = list(set(mobiles))
        for idx, number in enumerate(mobiles):
            mobile_numbers[idx] = number
        return mobile_numbers

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
        data = {}
        try:
            if self.extension == CVReader.__doc:
                data = self.read_doc()
            elif self.extension == CVReader.__docx:
                data = self.read_docx()
            elif self.extension == CVReader.__pdf:
                data = self.read_pdf()
            else:
                self.skip = True
            self.data = data
        except Exception as err:
            log.debug(traceback.format_exc())
            log.error("Error reading the file: {}. Please contact developer to fix it.".format(self.filename))
        finally:
            return self


class CVManager(object):

    output_folder = os.path.join(os.getcwd(), "output")
    path_txt_files = os.path.join(output_folder, "txts")
    path_doc2docx_files = os.path.join(output_folder, "doc2docx")
    path_unread_files = os.path.join(output_folder, "resumes_unread")
    path_unread_debug_files = os.path.join(path_unread_files, "debug")

    def __init__(self, args):
        self.args = args
        self.data = []
        self.data_unread = []
        self.stats = {
            "total": [],
            "read": [],
            "unread": [],
            "skip": []
        }

    @classmethod
    def filename(cls, path):
        return os.path.split(path)[-1]

    @classmethod
    def setup(cls):
        os.makedirs(cls.path_txt_files, exist_ok=True)
        os.makedirs(cls.path_doc2docx_files, exist_ok=True)
        os.makedirs(cls.path_unread_files, exist_ok=True)
        os.makedirs(cls.path_unread_debug_files, exist_ok=True)

    @classmethod
    def cleanup(cls):
        shutil.rmtree(CVManager.path_txt_files, ignore_errors=True)
        shutil.rmtree(CVManager.path_doc2docx_files, ignore_errors=True)

    @classmethod
    def valid(cls, data):
        if not data["first_name"]:
            return False
        if not data["mobile1"]:
            return False
        if not data["email1"]:
            return False
        return True

    def get(self):
        path = os.path.join(self.args.destination, "*")
        paths_all = glob.glob(path, recursive=True)
        paths = list(filter(lambda x: "~$" not in x, paths_all))

        self.stats["total"].extend(paths)
        for path in paths:
            reader = CVReader(path).read()
            if reader.skip:
                self.stats["skip"].append(path)
                continue

            data = reader.data
            log.debug("Fetched data: \n{}".format(json.dumps(data, indent=4)))
            if data:
                if self.valid(data):
                    self.data.append(data)
                    self.stats["read"].append(path)
                else:
                    self.data_unread.append(data)
                    self.stats["unread"].append(path)

                    src, dest = path, os.path.join(CVManager.path_unread_files, self.filename(path))
                    shutil.copyfile(src, dest)

                    src = reader.text_path
                    dest = os.path.join(CVManager.path_unread_debug_files, self.filename(reader.text_path))
                    shutil.copyfile(src, dest)
        return self

    def save(self):
        df = pd.DataFrame(self.data)
        path = os.path.join(CVManager.output_folder, "cv_info.xlsx")
        df.to_excel(path, index=False)

        df = pd.DataFrame(self.data_unread)
        path = os.path.join(CVManager.output_folder, "cv_unread_info.xlsx")
        df.to_excel(path, index=False)

    def conclude(self):
        if self.stats["skip"]:
            log.info("Skipping below files, since the document is not in any of the supported format: {}".format(
                ", ".join(CVReader.docs_supported)))
            for path in self.stats["skip"]:
                log.info("Skipped: {}".format(path))

        total, read, unread, skip = self.stats["total"], self.stats["read"], self.stats["unread"], self.stats["skip"]
        percent = round((len(read) / (len(total) - len(skip))) * 100, 2)
        log.info("""File read stats can be found below: 
        
        Total files attempted to read: {} 
        Total files skipped: {}
        Total files read successfully: {}
        Total files unread: {}
        Success: {} %
        """.format(len(total), len(skip), len(read), len(unread), percent))


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
    manager = CVManager(args)
    try:
        manager.get()
    except BaseException as err:
        log.error(err)
    finally:
        manager.save()
    manager.conclude()
    CVManager.cleanup()

    end = dt.now().strftime("%d-%m-%Y %H:%M:%S %p")
    log.info("Script ends at: {}".format(end))


if __name__ == "__main__":
    main()
