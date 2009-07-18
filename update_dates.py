from appscript import *
import re

DATE_FORMATS = (('%m/%d/%y',  r'\d{1,2}/\d{1,2}/\d{1,2}'       ), # T-Mobile
                ('%m.%d.%y',  r'\d{1,2}\.\d{1,2}\.\d{1,2}'     ), # iFixit
                ('%b %d, %Y', r'[A-Z][a-z][a-z] \d{1,2}, \d{4}'), # AmerenIP
                ('%B %d, %Y', r'[A-Z][a-z]+  ?\d{1,2}, ?\d{4}' ), # Amazon
                ('of %Y%m%d', r'of \d{8}'                      ), # Amazon
                ('%m/%d/%Y',  r'\d{1,2}/\d{1,2}/\d{4}'         ), # Busey
                ('%b %d %Y',  r'[A-Z]{3} \d{1,2} \d{4}'        ), # State Farm
                ('%d %b %Y',  r'\d{1,2} [A-Z][A-Za-z]{2} \d{4}'), # Apple
                ('%Y-%m-%d',  r'\d{4}-\d{2}-\d{2}'             ), # MacSpeech
                ('%Y-%m',     r'\d{4}-\d{2}'                   ), # filename
                ('%m1%d/%y',  r'\d{1,2}1\d{1,2}/\d{1,2}'       ), # T-Mo bad OCR
                ('%m/%d1%y',  r'\d{1,2}/\d{1,2}1\d{1,2}'       ), # T-Mo bad OCR
                ('%m/%d/%y',  r'\d{1,2}/ \d{1,2}/ \d{1,2}'     ), # T-Mo bad OCR
                ('%m/%d/%Y',
                 r'(?:\d ?){1,2}/ (?:\d ?){1,2}/ (?:\d ?){4}'  ), # Busey bad OCR
                )

RE_DATE = re.compile('|'.join(r'(\b%s\b)' % regex
                              for format, regex in DATE_FORMATS))

def extract_date(contents, match=None):
    no_match = []
    for m in RE_DATE.finditer(contents):
        matched_format = m.lastindex
        format = DATE_FORMATS[matched_format - 1][0]
        # note: spaces in strptime format match zero or more spaces, this is OK
        matched = m.group(matched_format).replace(' ', '')
        try:
            parsed = datetime.datetime.strptime(matched, format)
        except ValueError, e: # not a date
            no_match.append((matched, format, e))
            continue
        if not match or (match.year, match.month) == (parsed.year, parsed.month):
            return parsed.date(), no_match
        no_match.append(m.group(matched_format))
    return None, no_match

EagleFiler = app(id='com.c-command.EagleFiler')
Paper = EagleFiler.documents['Paper.eflibrary']

for record in Paper.library_records[its.kind=='PDF']():
    title = record.title()
    hint, no_match = extract_date(title)

    contents = record.contents()
    extracted, no_match = extract_date(contents, hint)

    if not extracted:
        print title, hint
        for nm in no_match:
            print '  no match', nm
        if not hint:
            continue

    record.creation_date.set(extracted or hint)
