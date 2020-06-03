"""Tools for writing RTF documents"""

class RtfDocument:
    """Create an RTF document and write content to it"""

    def __init__(self, monospace=False, header_text=None, first_header_text=None):
        """Start a new document from the template

        Arguments:
         - monospace (boolean): True to use Courier New throughout the doc; False to use
            Times New Roman
         - header_text (str): Text to add to the header for every page. If header_text
            includes the string "_P#_", this will be replaced with a page number field.
         - first_header_text (str): If not None, will use a different header for the
            first page; otherwise, same as header_text
        """

        self.document = TEMPLATE
        if monospace:
            self.document = self.document.replace("__STYLESHEET__", COURIER_STYLESHEET)
            self.font = 'f3'
        else:
            self.document = self.document.replace("__STYLESHEET__", TIMES_STYLESHEET)
            self.font = 'f0'
        if header_text is not None:
            self.insert_rtf(self.new_header(header_text))
        if first_header_text is not None:
            self.insert_rtf(self.new_header(first_header_text, first_page=True))
            self.document = self.document.replace("__TITLEPG__", r'\titlepg')
        else:
            self.document = self.document.replace("__TITLEPG__", '')

    def insert_rtf(self, new_content):
        """Insert new content at the 'head'"""

        self.document = self.document.replace("__HEAD__", f"{new_content}\n__HEAD__")

    def new_header(self, text, first_page=False):
        """Build a new RTF header content block.

        Arguments:
         - text (str): Text to appear in the header, align right. "_P#_" will be
            replace with a page number field.
         - first_page (str): If true, make a first page header; otherwise, a
            general header

        Returns (str): The RTF code for the header
        """
        text = text.replace("_P#_", r"{\field{\*\fldinst{PAGE}}{\fldrslt}}")
        out = HEADER.replace("__TEXT__", text)
        out = out.replace("__FONT__", self.font)
        if first_page:
            return out.replace("__HEADERTYPE__", "headerf")
        else:
            return out.replace("__HEADERTYPE__", "header")

    def add_lines(self, lines, style="normal", indent=True, double_space=True, new_page=False):
        """Add one or more lines of a particular style to the end of the document.

        Arguments:
         - lines (list of str): The lines of text to append
         - style (str): The style to apply. One of:
            - normal
            - centered
            - h1
            - h2
            - title
            - subtitle
         - indent (boolean): If True, apply a first-line half-inch indent. Ignored
            if style is not "normal".
         - double_space (boolean): If True, double space after each line
         - new_page (boolean): If True, begin a new page before appending lines
        """

        if style == "normal":
            if indent:
                styling = NORMAL_STYLE.replace("__INDENT__", "720")
            else:
                styling = NORMAL_STYLE.replace("__INDENT__", "0")
        else:
            if style == "h1":
                scode = "s1"
            elif style == "h2":
                scode = "s2"
            elif style == "centered":
                scode = "s3"
            elif style == "title":
                scode = "s15"
            elif style == "subtitle":
                scode = "s16"
            styling = HEADING_STYLE.replace("__STYLE__", scode)
        if double_space:
            styling = styling.replace("__SPACING__", "480")
        else:
            styling = styling.replace("__SPACING__", "240")
        styling = styling.replace("__FONT__", self.font)
        if new_page:
            # Only apply page break to first line as a separate style
            styling = styling.replace("__PAGEBR__", r"\pagebb")
            self.insert_rtf(styling)
            line = lines[0].replace("\n", r" \par ")
            line = line.replace("\t", r" \tab ")
            self.insert_rtf("{" + line + r" \par}")
            self.insert_rtf(styling.replace("\pagebb", ""))
            for line in lines[1:]:
                line = line.replace("\n", r" \par ")
                line = line.replace("\t", r" \tab ")
                self.insert_rtf("{" + line + r" \par}")
        else:
            styling = styling.replace("__PAGEBR__", "")
            self.insert_rtf(styling)
            for line in lines:
                line = line.replace("\n", r" \par ")
                line = line.replace("\t", r" \tab ")
                self.insert_rtf("{" + line + r" \par}")

    def dump(self):
        """Get a str of the RTF document"""

        return self.document.replace("__HEAD__", "")


# ----Constants containing RTF code/templates----

TEMPLATE = r"""{\rtf1\ansi\ansicpg1252\uc0\stshfdbch0\stshfloch0\stshfhich0\stshfbi0\deff0\adeff0
{\fonttbl{\f0\froman\fcharset0\fprq2 Times New Roman;}{\f1\froman\fcharset2\fprq2 Symbol;}
{\f2\fswiss\fcharset0\fprq2 Arial;}{\f3\fnil\fcharset0 Courier New;}}
{\colortbl;\red0\green0\blue0;\red102\green102\blue102 ;}{\stylesheet__STYLESHEET__}
{\*\rsidtbl\rsid10976062}{\mmathPr\mbrkBin0\mbrkBinSub0\mdefJc1\mdispDef1\minterSp0\mintLim0
\mintraSp0\mlMargin0\mmathFont0\mnaryLim1\mpostSp0\mpreSp0\mrMargin0\msmallFrac0\mwrapIndent1440
\mwrapRight0}
\deflang1033\deflangfe2052\adeflang1025\jexpand\showxmlerrors1\validatexml1{\*\wgrffmtfilter 013f}
\viewkind1\viewscale100\fet0\ftnbj\aenddoc\ftnrstcont\aftnrstcont\ftnnar\aftnnrlc\widowctrl
\nospaceforul\nolnhtadjtbl\alntblind\lyttblrtgr\dntblnsbdb\noxlattoyen\wrppunct\nobrkwrptbl
\expshrtn\snaptogridincell\asianbrkrule\htmautsp\noultrlspc\useltbaln\splytwnine\ftnlytwnine
\lytcalctblwd\allowfieldendsel\lnbrkrule\nouicompat\nofeaturethrottle1\formshade\nojkernpunct
\dghspace180\dgvspace180\dghorigin1800\dgvorigin1440\dghshow1\dgvshow1\dgmargin\pgbrdrhead
\pgbrdrfoot\sectd__TITLEPG__\sectlinegrid360\pgwsxn12240\pghsxn15840\marglsxn1440\margrsxn1440
\margtsxn1440\margbsxn1440\guttersxn0\headery708\footery708\colsx708\ltrsect
__HEAD__
{\*\latentstyles\lsdstimax267\lsdlockeddef0\lsdsemihiddendef0\lsdunhideuseddef0
\lsdqformatdef0\lsdprioritydef0{\lsdlockedexcept\lsdqformat1 Normal;
\lsdqformat1 heading 1;\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 2;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 3;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 4;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 5;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 6;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 7;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 8;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 heading 9;
\lsdsemihidden1\lsdunhideused1\lsdqformat1 caption;\lsdqformat1 Title;\lsdqformat1 Subtitle;
\lsdqformat1 Strong;\lsdqformat1 Emphasis;\lsdsemihidden1\lsdpriority99 Placeholder Text;
\lsdqformat1\lsdpriority1 No Spacing;\lsdpriority60 Light Shading;\lsdpriority61 Light List;
\lsdpriority62 Light Grid;\lsdpriority63 Medium Shading 1;\lsdpriority64 Medium Shading 2;
\lsdpriority65 Medium List 1;\lsdpriority66 Medium List 2;\lsdpriority67 Medium Grid 1;
\lsdpriority68 Medium Grid 2;\lsdpriority69 Medium Grid 3;\lsdpriority70 Dark List;
\lsdpriority71 Colorful Shading;\lsdpriority72 Colorful List;\lsdpriority73 Colorful Grid;
\lsdpriority60 Light Shading Accent 1;\lsdpriority61 Light List Accent 1;
\lsdpriority62 Light Grid Accent 1;\lsdpriority63 Medium Shading 1 Accent 1;
\lsdpriority64 Medium Shading 2 Accent 1;\lsdpriority65 Medium List 1 Accent 1;
\lsdsemihidden1\lsdpriority99 Revision;\lsdqformat1\lsdpriority34 List Paragraph;
\lsdqformat1\lsdpriority29 Quote;\lsdqformat1\lsdpriority30 Intense Quote;
\lsdpriority66 Medium List 2 Accent 1;\lsdpriority67 Medium Grid 1 Accent 1;
\lsdpriority68 Medium Grid 2 Accent 1;\lsdpriority69 Medium Grid 3 Accent 1;
\lsdpriority70 Dark List Accent 1;\lsdpriority71 Colorful Shading Accent 1;
\lsdpriority72 Colorful List Accent 1;\lsdpriority73 Colorful Grid Accent 1;
\lsdpriority60 Light Shading Accent 2;\lsdpriority61 Light List Accent 2;
\lsdpriority62 Light Grid Accent 2;\lsdpriority63 Medium Shading 1 Accent 2;
\lsdpriority64 Medium Shading 2 Accent 2;\lsdpriority65 Medium List 1 Accent 2;
\lsdpriority66 Medium List 2 Accent 2;\lsdpriority67 Medium Grid 1 Accent 2;
\lsdpriority68 Medium Grid 2 Accent 2;\lsdpriority69 Medium Grid 3 Accent 2;
\lsdpriority70 Dark List Accent 2;\lsdpriority71 Colorful Shading Accent 2;
\lsdpriority72 Colorful List Accent 2;\lsdpriority73 Colorful Grid Accent 2;
\lsdpriority60 Light Shading Accent 3;\lsdpriority61 Light List Accent 3;
\lsdpriority62 Light Grid Accent 3;\lsdpriority63 Medium Shading 1 Accent 3;
\lsdpriority64 Medium Shading 2 Accent 3;\lsdpriority65 Medium List 1 Accent 3;
\lsdpriority66 Medium List 2 Accent 3;\lsdpriority67 Medium Grid 1 Accent 3;
\lsdpriority68 Medium Grid 2 Accent 3;\lsdpriority69 Medium Grid 3 Accent 3;
\lsdpriority70 Dark List Accent 3;\lsdpriority71 Colorful Shading Accent 3;
\lsdpriority72 Colorful List Accent 3;\lsdpriority73 Colorful Grid Accent 3;
\lsdpriority60 Light Shading Accent 4;\lsdpriority61 Light List Accent 4;
\lsdpriority62 Light Grid Accent 4;\lsdpriority63 Medium Shading 1 Accent 4;
\lsdpriority64 Medium Shading 2 Accent 4;\lsdpriority65 Medium List 1 Accent 4;
\lsdpriority66 Medium List 2 Accent 4;\lsdpriority67 Medium Grid 1 Accent 4;
\lsdpriority68 Medium Grid 2 Accent 4;\lsdpriority69 Medium Grid 3 Accent 4;
\lsdpriority70 Dark List Accent 4;\lsdpriority71 Colorful Shading Accent 4;
\lsdpriority72 Colorful List Accent 4;\lsdpriority73 Colorful Grid Accent 4;
\lsdpriority60 Light Shading Accent 5;\lsdpriority61 Light List Accent 5;
\lsdpriority62 Light Grid Accent 5;\lsdpriority63 Medium Shading 1 Accent 5;
\lsdpriority64 Medium Shading 2 Accent 5;\lsdpriority65 Medium List 1 Accent 5;
\lsdpriority66 Medium List 2 Accent 5;\lsdpriority67 Medium Grid 1 Accent 5;
\lsdpriority68 Medium Grid 2 Accent 5;\lsdpriority69 Medium Grid 3 Accent 5;
\lsdpriority70 Dark List Accent 5;\lsdpriority71 Colorful Shading Accent 5;
\lsdpriority72 Colorful List Accent 5;\lsdpriority73 Colorful Grid Accent 5;
\lsdpriority60 Light Shading Accent 6;\lsdpriority61 Light List Accent 6;
\lsdpriority62 Light Grid Accent 6;\lsdpriority63 Medium Shading 1 Accent 6;
\lsdpriority64 Medium Shading 2 Accent 6;\lsdpriority65 Medium List 1 Accent 6;
\lsdpriority66 Medium List 2 Accent 6;\lsdpriority67 Medium Grid 1 Accent 6;
\lsdpriority68 Medium Grid 2 Accent 6;\lsdpriority69 Medium Grid 3 Accent 6;
\lsdpriority70 Dark List Accent 6;\lsdpriority71 Colorful Shading Accent 6;
\lsdpriority72 Colorful List Accent 6;\lsdpriority73 Colorful Grid Accent 6;
\lsdqformat1\lsdpriority19 Subtle Emphasis;\lsdqformat1\lsdpriority21 Intense Emphasis;
\lsdqformat1\lsdpriority31 Subtle Reference;\lsdqformat1\lsdpriority32 Intense Reference;
\lsdqformat1\lsdpriority33 Book Title;\lsdsemihidden1\lsdunhideused1\lsdpriority37 Bibliography;
\lsdsemihidden1\lsdunhideused1\lsdqformat1\lsdpriority39 TOC Heading;}}}"""

TIMES_STYLESHEET = r"""{\s0\snext0\sqformat\spriority0\fi720\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\ql\faauto\sl480\slmult1\rtlch\ab0\ai0
\af0\afs24\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 Normal;}
{\s1\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af0\afs24
\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 heading 1;}
{\s2\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af0\afs24
\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 heading 2;}
{\s3\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af0\afs24
\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 centered;}
{\s4\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qr\faauto\s1240\slmult1\rtlch\ab0\ai0\af0\afs24
\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 page header;}
{\*\cs10\additive\ssemihidden\spriority0 Default Paragraph Font;}
{\s15\sbasedon0\snext15\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af0\afs24
\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 Title;}
{\s16\sbasedon0\snext16\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af0\afs24
\ltrch\b0\i0\fs24\f0\strike0\ulnone\cf1 Subtitle;}"""

COURIER_STYLESHEET = r"""{\s0\snext0\sqformat\spriority0\fi720\sb0\sa0\aspalpha\aspnum\adjustright
\widctlpar\ltrpar\li0\lin0\ri0\rin0\ql\faauto\sl480\slmult1\rtlch\ab0\ai0\af3\afs24\ltrch\b0\i0
\fs24\loch\af3\dbch\af3\hich\f3\strike0\ulnone\cf1 Normal;}
{\s1\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af3\afs24
\ltrch\b0\i0\fs24\loch\af3\dbch\af3\hich\f3\strike0\ulnone\cf1 heading 1;}
{\s2\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af3\afs24
\ltrch\b0\i0\fs24\loch\af3\dbch\af3\hich\f3\strike0\ulnone\cf1 heading 2;}
{\s3\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af3\afs24
\ltrch\b0\i0\fs24\loch\af3\dbch\af3\hich\f3\strike0\ulnone\cf1 centered;}
{\s4\sbasedon0\snext0\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qr\faauto\s1240\slmult1\rtlch\ab0\ai0\af3\afs24
\ltrch\b0\i0\fs24\f3\strike0\ulnone\cf1 page header;}
{\*\cs10\additive\ssemihidden\spriority0 Default Paragraph Font;}
{\s15\sbasedon0\snext15\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af3\afs24
\ltrch\b0\i0\fs24\loch\af3\dbch\af3\hich\f3\strike0\ulnone\cf1 Title ;}
{\s16\sbasedon0\snext16\styrsid15694742\sqformat\spriority0\keep\keepn\fi0\sb0\sa0\aspalpha\aspnum
\adjustright\widctlpar\ltrpar\li0\lin0\ri0\rin0\qc\faauto\sl480\slmult1\rtlch\ab0\ai0\af3\afs24
\ltrch\b0\i0\fs24\loch\af3\dbch\af3\hich\f3\strike0\ulnone\cf1 Subtitle;}"""

HEADER = r"""{\__HEADERTYPE__\pard\plain\itap0\ilvl0\s4\fi0\sb0\sa0\aspalpha\aspnum\adjustright
\brdrt\brdrl\brdrb\brdrr\brdrbtw\brdrbar\widctlpar\ltrpar\li0\lin0\ri0\rin0\qr\faauto\s1240
\slmult1\rtlch\ab0\ai0\a__FONT__\afs24\ltrch\b0\i0\fs24\__FONT__\strike0\ulnone\cf1{__TEXT__}}"""

NORMAL_STYLE = r"""\pard\plain\itap0\s0__PAGEBR__\ilvl0\tqr\tx9360\fi__INDENT__\sb0\sa0
\aspalpha\aspnum\adjustright\brdrt\brdrl\brdrb\brdrr\brdrbtw\brdrbar\widctlpar\ltrpar\li0\lin0\ri0
\rin0\ql\faauto\sl__SPACING__\slmult1\rtlch\ab0\ai0\a__FONT__\afs24\ltrch\b0\i0\fs24\__FONT__
\strike0\ulnone\cf1"""

HEADING_STYLE = r"""\pard\plain\itap0\__STYLE____PAGEBR__\keep\keepn\ilvl0\fi0\sb0\sa0\aspalpha
\aspnum\adjustright\brdrt\brdrl\brdrb\brdrr\brdrbtw\brdrbar\widctlpar\ltrpar\li0\lin0\ri0\rin0
\qc\faauto\sl__SPACING__\slmult1\rtlch\ab0\ai0\a__FONT__\afs24\ltrch\b0\i0\fs24\__FONT__\strike0
\ulnone\cf1"""
