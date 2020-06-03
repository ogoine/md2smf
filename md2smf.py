
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

SINGLE_LEFT_QUOTE = r'\lquote '
SINGLE_RIGHT_QUOTE = r'\rquote '
DOUBLE_LEFT_QUOTE = r'\ldblquote '
DOUBLE_RIGHT_QUOTE = r'\rdblquote '
EM_DASH = r'\emdash '

def trunc(number, places):
    """Truncate the given number to the nearest 10^places"""

    return 10 ** places * int(number / 10 ** places)

def count_words(lines):
    """Count words in a list of lines, excluding headings and comments"""

    body = [i for i in lines if i[0] not in ("#", ">")]
    return len(' '.join(body).split(' '))

def heading_level(line):
    """Return heading level of line (1, 2, 3, or 0 for normal)"""

    for i in range(4):
        if line[i] != '#':
            return i
    return 3

def char_weight(string, index):
    """Return a weight of 0-2 for the char at the given index, used for deciding quote direction"""

    if index < 0 or index >= len(string):
        return 0
    if string[index] == ' ':
        return 0
    if string[index].isalnum():
        return 2
    return 1 # symbols

def substitute(string, index, substitute):
    """Replace the char at index in string with substitute (str)"""

    return string[:index] + substitute + string[index + 1:]

def smart_replace(line):
    """Replace dumb quotes with smart, apply italics, etc."""

    # start with simple substitutions
    line = line.replace("--", EM_DASH)
    line = line.replace('  ', ' ')
    line = line.replace('  ', ' ')

    # now parse more contextual ones
    italics_on = False
    index = 0
    while index < len(line):
        if line[index] == "'":
            if char_weight(line, index - 1) < char_weight(line, index + 1):
                line = substitute(line, index, SINGLE_LEFT_QUOTE)
                index += len(SINGLE_LEFT_QUOTE)
            else:
                line = substitute(line, index, SINGLE_RIGHT_QUOTE)
                index += len(SINGLE_RIGHT_QUOTE)
        elif line[index] == '"':
            if char_weight(line, index - 1) < char_weight(line, index + 1):
                line = substitute(line, index, DOUBLE_LEFT_QUOTE)
                index += len(DOUBLE_LEFT_QUOTE)
            else:
                line = substitute(line, index, DOUBLE_RIGHT_QUOTE)
                index += len(DOUBLE_RIGHT_QUOTE)
        elif line[index] == '*':
            if italics_on:
                line = substitute(line, index, r'\i0 ')
                index += len(r'\i0 ')
            else:
                line = substitute(line, index, r'\i ')
                index += len(r'\i ')
            italics_on = not italics_on
        else:
            index += 1
    if italics_on:
        line += r'\i0 '
    return line

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
                self.parse_heading()
                new_division = True
            elif new_level == 3:
                self.parse_section(hide=new_division)
            else:
                self.parse_normal()
                new_division = False
        return self.doc

    def parse_heading(self):
        """Parse one or more part or chapter headers, adding to self.doc"""

        already_broken = False
        level = heading_level(self.lines[0])
        while self.lines and level in (1, 2):
            line = self.lines[0][level:].strip()
            blanks = [''] if already_broken else [''] * 8
            self.doc.add_lines(blanks, new_page=not already_broken)
            already_broken = True
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
            level = heading_level(self.lines[0])

    def parse_section(self, hide=False):
        """Parse one or more section headers, adding to self.doc"""

        while self.lines and heading_level(self.lines[0]) == 3:
            if not hide:
                self.doc.add_lines(['#'], style='centered')
            self.lines.pop(0)

    def parse_normal(self):
        """Parse self.lines from the normal-level, adding to self.doc"""

        normal_lines = []
        while self.lines and heading_level(self.lines[0]) == 0:
            if self.monospace:
                normal_lines.append(self.lines[0])
            else:
                normal_lines.append(smart_replace(self.lines[0]))
            self.lines.pop(0)
        self.doc.add_lines(normal_lines)

def convert_to_smf(md_document, head_file=None, monospace=False, name_chapters=False, name_parts=False):

    # read info
    lines = [i.strip() for i in md_document.splitlines() if i.strip() and i.strip()[0] != '>']
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
    if word_count < 20000:
        head[0] = f'{head[0]}\tabout {round(word_count, 2):,} words'
    for i in range(18 - len(head)):
        head.append('')
    doc.add_lines(head, indent=False, double_space=False)
    doc.add_lines([title], style="title")
    doc.add_lines([f'By {author}'], style="subtitle")
    if word_count >= 20000:
        doc.add_lines([''] * 9 + [f'about {round(word_count, 3):,} words'], style='centered')
    else:
        doc.add_lines(['']) 

    # write document body
    return MdParser(doc, input, monospace, name_chapters, name_parts).parse().dump()

def print_word_count(md_document):
    lines = [i.strip() for i in md_document.splitlines() if i.strip()]
    print(f'Word count: {count_words(lines)}')
    todos = len([i for i in lines if i.startswith('> TODO:')])
    print(f'Remaining to-dos: {todos}')

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--head-file', help='File containing manuscript head information')
    parser.add_argument('-m', '--monospace', help='Use a monospace font', action='store_true')
    parser.add_argument('-c', '--chapter-names', help='Display chapter names in addition to numbering', action='store_true')
    parser.add_argument('-p', '--part-names', help='Display part names in addition to numbering', action='store_true')
    parser.add_argument('-w', '--word-count', help='Display word count information and exit', action='store_true')
    args = parser.parse_args()
    if args.word_count:
        print_word_count(sys.stdin.read())
    else:
        print(convert_to_smf(
            sys.stdin.read(),
            args.head_file,
            args.monospace,
            args.chapter_names,
            args.part_names))

if __name__ == "__main__":
    main()
