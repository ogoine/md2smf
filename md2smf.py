
import argparse
import sys

from rtf_builder import RtfDocument

FAKE_HEAD = ['Your Name',
             'Address Line 1',
             'Address Line 2',
             'Address Line 3',
             'City, ST Zip',
             'email@example.com',
             '(555) 555 - 5555']

def count_words(lines):
    """Count words in a list of lines, excluding headings and comments"""

    body = [i for i in lines if i[0] not in ("#", ">")]
    return 100 * int(len(' '.join(body).split(' ')) / 100)

def heading_level(line):
    """Return heading level of line (1, 2, 3, or 0 for normal)"""

    for i in range(4):
        if line[i] != '#':
            return i
    return 3


class MdParser:
    """Parse Markdown document body and build out an initialized RtfDocument from it"""

    def __init__(self, rtf_doc, md_lines, monospace=False, name_chapters=False, name_parts=False):
        self.doc = rtf_doc
        self.lines = md_lines
        self.monospace = monospace
        self.name_chapters = name_chapters
        self.name_parts = name_parts
        self.part_count = 0
        self.chapter_count = 0

    def parse(self):
        """Parse all lines and add them to self.doc"""

        # track if we are starting a new paragraph or chapter
        # if so, hide the first sections until we encounter some normal text
        new_division = True
        while self.lines:
            new_level = heading_level(self.lines[0])
            if new_level in (1, 2):
                self.parse_heading(new_level)
                new_division = True
            elif new_level == 3:
                self.parse_section(hide=new_division)
            else:
                if self.lines[0][0] == '>':
                    # don't assume a normal section is beginning if it's only a comment
                    self.lines.pop(0)
                else:
                    self.parse_normal()
                    new_division = False
        return self.doc

    def parse_heading(self, level):
        """Parse one or more part or chapter headers, adding to self.doc"""

        while self.lines and heading_level(self.lines[0]) == level:
            line = self.lines[0][level:].strip()
            blanks = ['' for i in range(8)]
            self.doc.add_lines(blanks, new_page=True)
            if level == 1:
                self.part_count += 1
                text = f'Part {self.part_count}'
                if self.name_parts:
                    text += f': {line}'
            elif level == 2:
                self.chapter_count += 1
                text = f'Chapter {self.chapter_count}'
                if self.name_chapters:
                    text += f': {line}'
            self.doc.add_lines([text], style=f'h{level}')
            self.doc.add_lines(['']) 
            self.lines.pop(0)

    def parse_section(self, hide=False):
        """Parse one or more section headers, adding to self.doc"""

        while self.lines and heading_level(self.lines[0]) == 3:
            if not hide:
                self.doc.add_lines(['#'], style='h3')
            self.lines.pop(0)

    def parse_normal(self):
        """Parse self.lines from the normal-level, adding to self.doc"""

        normal_lines = []
        while self.lines and heading_level(self.lines[0]) == 0:
            if self.lines[0][0] != '>':
                normal_lines.append(self.lines[0])
            self.lines.pop(0)
        self.doc.add_lines(normal_lines)

def convert_to_smf(md_document, head_file=None, monospace=False, name_chapters=False, name_parts=False):

    # read info
    lines = [i.strip() for i in md_document.splitlines() if i.strip()]
    word_count = count_words(lines)
    # NEED TO REVERSE SORT
    input = lines.copy()
    if input[0] != "# HEAD":
        raise RuntimeError(f"First line of document must be '# HEAD', not '{input[0]}'")
    input.pop(0)
    if input[0] == '#':
        raise RuntimeError(f'Missing manuscript title.')
    title = input[0]
    input.pop(0)
    if input[0] == '#':
        raise RuntimeError(f'Missing manuscript short name.')
    short_name = input[0]
    input.pop(0)
    if input[0] == '#':
        raise RuntimeError(f'Missing author name.')
    author = input[0]
    input.pop(0)
    last = author.split(' ')[-1]
    if head_file:
        with open(head_file) as fileobj:
            # Only keep as many lines as you see in the FAKE_HEAD above
            head = fileobj.read().splitlines()[:len(FAKE_HEAD)]
    else:
        head = FAKE_HEAD

    # write document head
    doc = RtfDocument(
        monospace,
        header_text=f'{last} / {short_name} / _P#_',
        first_header_text='')
    head[0] = f'{head[0]}\tAbout {word_count} words'
    for i in range(18 - len(head)):
        head.append('')
    doc.add_lines(head, indent=False, double_space=False)
    doc.add_lines([title], style="title")
    doc.add_lines([f'by {author}'], style="subtitle")
    doc.add_lines(['']) 

    # write document body
    return MdParser(doc, input, monospace, name_chapters, name_parts).parse().dump()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--head-file', help='File containing manuscript head information')
    parser.add_argument('-m', '--monospace', help='Use a monospace font', action='store_true')
    parser.add_argument('-c', '--chapter-names', help='Display chapter names in addition to numbering', action='store_true')
    parser.add_argument('-p', '--part-names', help='Display part names in addition to numbering', action='store_true')
    args = parser.parse_args()
    print(convert_to_smf(
        sys.stdin.read(),
        args.head_file,
        args.monospace,
        args.chapter_names,
        args.part_names))

if __name__ == "__main__":
    main()
