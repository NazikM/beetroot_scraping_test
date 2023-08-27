import pdfplumber
import re
import math

from excel import ExcelWriter


class PublicationExtractor:
    def __init__(self, file_name):
        self.file_name = file_name
        self.result = []
        self.default_publication = {
            'p': '',
            'title': '',
            'author_info': '',
            'organization': '',
            'location': '',
            'content': ''
        }

    @staticmethod
    def extract_location(publication):
        location = re.findall(r'[A-Z][a-z]+,\s[A-Z][a-z]+', publication['organization'])
        publication['location'] = '; '.join(location)

    def append_publication(self, publication):
        """Append publication to result list and creates new from default"""
        self.extract_location(publication)
        self.result.append(publication)

    def process_pages(self):
        with pdfplumber.open(file_name) as pdf:
            # Using to understand what we're extracting right now 0 (title) -> 1 (author info) -> 2 (main content)
            stage = 0
            publication = self.default_publication.copy()
            for page_number in range(6, 63):
                page = pdf.pages[page_number]

                # Split every page to 2 pages for easier scraping
                left_page_side = page.crop((0, 43, page.width / 2, 730))
                right_page_side = page.crop((page.width / 2, 43, page.width, 730))

                for side in (left_page_side, right_page_side):
                    for char in side.chars:
                        # PXXX is written with size 9.5 everywhere
                        if math.isclose(char['size'], 9.5) and 'Bold' in char['fontname']:
                            if char['text'] == " ":
                                continue
                            if len(publication['p']) == 4:
                                self.append_publication(publication)
                                publication = self.default_publication.copy()
                            publication['p'] += char['text']
                            stage = 0
                        elif char['size'] in (8, 9):
                            if 'Bold' in char['fontname']:
                                publication['title'] += char['text']
                            elif 'Italic' in char['fontname']:
                                if stage == 2 or stage == 1 and char['size'] != 8:
                                    publication['content'] += char['text']
                                    continue

                                if char['size'] == 8:
                                    publication['organization'] += char['text']
                                    stage = 1

                                if stage == 0:
                                    publication['author_info'] += char['text']
                            else:
                                if stage == 1:
                                    stage = 2
                                publication['content'] += char['text']
            self.append_publication(publication)

    def save_to_excel(self):
        excel_writer = ExcelWriter("result.xlsx")
        data_to_write = [
            (row['author_info'], row['organization'], row['location'], row['p'], row['title'], row['content']) for row
            in self.result]
        excel_writer.write_data(data_to_write)
        excel_writer.save()

    def extract_publications(self):
        self.process_pages()
        self.save_to_excel()


if __name__ == "__main__":
    file_name = "Abstract Book from the 5th World Psoriasis and Psoriatic Arthritis Conference 2018.pdf"
    extractor = PublicationExtractor(file_name)
    extractor.extract_publications()
