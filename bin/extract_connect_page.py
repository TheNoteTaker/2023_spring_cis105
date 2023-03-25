from bs4 import BeautifulSoup
import re


# TODO: Add class to extract the notes page from the iframe of the reader page
class NotesPage:
    # BOOK_URL = r"https://epub-factory-cdn.mheducation.com/publish/sn_7bac8/7/1080mp4/OPS/s9ml/chapter10/reader_02.xhtml"
    # DICT_URL = r"https://epub-factory-cdn.mheducation.com/publish/sn_7bac8/7/1080mp4/OPS/s9ml/glossary.xhtml"
    BASE_RE = r"http[s]?://(?:www.)?epub-factory[^\./]*(?:\.[^/]+)+(?:/[^/]+)+"
    EXT_RE = ".xhtml"
    DICT_RE = re.compile(f"{BASE_RE}/glossary{EXT_RE}")
    CHAPTER_RE = \
        re.compile(fr"({BASE_RE}/(chapter)(\d+)/(reader.)(\d+){EXT_RE})")

    def __init__(self, html: str):
        self._soup = BeautifulSoup(html, 'xml')
        self.html = self._soup.prettify()

        # Placeholder text to substitute into urls
        self._page_text = "<PAGE_NUM>"
        self._chapter_text = "<CHAPTER_NUM>"

    def get_chapter_url_format(self) -> dict:
        """
        Gets the first chapter url result.

        Gets the first chapter url result found in `_soup` and isolates
        the chapter prefix (the "chapter" in chapter01), the page
        prefix (the "page_" in "page_01", and the length of the number
        sections of both the chapter and page sections.

        Returns:
            `dict` containing information about the url:
                url itself formatted to replace the page and chapter
                    numbers.
                The chapter prefix
                The chapter number length
                The page prefix
                The page number length
        """
        url, chapter_prefix, chapter_num, page_prefix, page_num = \
            self.CHAPTER_RE.search(self.html).groups()

        url = url.replace(page_num, self._page_text)
        url = url.replace(chapter_num, self._chapter_text)

        chapter_url_format = {
            "url_format": url,
            "chapter_format": chapter_prefix + chapter_num,
            "chapter_length": len(chapter_num),
            "page_format": page_prefix + page_num,
            "page_length": len(page_num)
        }

        return chapter_url_format

    def get_dict_url(self) -> str:
        """Returns the url where the dictionary is located."""
        return self.DICT_RE.search(self.html).group()

    def create_url(self, chapter_num: str, page_num: str) -> str:
        """
        Creates a valid url to go to a specific chapter and page number.

        Generates a valid McGraw Hill URL that will lead to chapter
        `chapter_num` and page `page_num`. Handles the padding of 0's
        for both the chapter number and page number to meet the minimum
        required.

        Args:
            chapter_num: Chapter to go to.
            page_num: Page to go to.

        Returns:
            Valid url to go to the selected chapter and page within the
            given book.
        """
        chapter_url_format = self.get_chapter_url_format()

        # Format both arguments to have correct 0 padding
        formatted_page_num = "0" * (
                chapter_url_format["page_length"] - len(page_num)
        ) + page_num

        formatted_chapter_num = "0" * (
                chapter_url_format["chapter_length"] - len(chapter_num)
        ) + chapter_num

        # Insert chapter and page numbers
        url = chapter_url_format["url_format"] \
            .replace(self._page_text, formatted_page_num)

        url = url.replace(self._chapter_text, formatted_chapter_num)

        return url

    def print_page(self) -> None:
        """Prints out prettified version of page html"""
        print(self._soup.html)

    def get_all_links(self) -> None:
        """Prints out all links in `_soup`"""
        for tag in self._soup.find_all("a"):
            print("Link:", tag["href"])


if __name__ == "__main__":
    with open("/home/rashino/Downloads/file.xhtml", 'r', encoding="utf-8") \
            as file:
        soup = NotesPage(file.read())

    ch_num = input("Please enter a chapter: ")
    pg_num = input("Please enter a page number: ")
    print(soup.create_url(ch_num, pg_num))
