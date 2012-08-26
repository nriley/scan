from appscript import *
from datetime import datetime
from osax import *
from plistlib import readPlist, writePlist
import os
import re
import time

PREFERENCES_PATH = \
    os.path.expanduser('~/Library/Preferences/net.sabi.UpdateDates.plist')

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
            parsed = datetime.strptime(matched, format)
        except ValueError, e: # not a date
            no_match.append((matched, format, e))
            continue
        if not match or (match.year, match.month) == (parsed.year, parsed.month):
            return parsed.date(), no_match
        no_match.append(m.group(matched_format))
    return None, no_match

def extract_source(title, hint):
    if hint:
        return title[:RE_DATE.search(title).start(0)].rstrip()
    else:
        return title

EagleFiler = app(id='com.c-command.EagleFiler')
Paper = EagleFiler.documents['Paper.eflibrary']

if not Paper.exists():
    EagleFiler.open(os.path.expanduser('~/Documents/Paper/Paper.eflibrary'))

def read_sources():
    return readPlist(PREFERENCES_PATH).get('Sources', [])

if os.path.exists(PREFERENCES_PATH):
    try:
        sources = read_sources()
    except:
        from subprocess import call
        call(['plutil', '-convert', 'xml1', PREFERENCES_PATH])
        sources = read_sources()
else:
    sources = []

def update_all():
    for record in Paper.library_records[its.kind=='PDF']():
        title = record.title()
        hint, no_match = extract_date(title)
        source = extract_source(title, hint)

        contents = record.contents()
        if re.search(re.escape(source), contents, re.IGNORECASE):
            if source in sources:
                sources.remove(source)
            sources.append(source)

        extracted, no_match = extract_date(contents, hint)

        if not extracted:
            print title, hint
            for nm in no_match:
                print '  no match', nm
            if not hint:
                continue

        record.creation_date.set(extracted or hint)

    sources.reverse() # most recently referenced ones at top

def scan_one():
    Acrobat = app(id='com.adobe.Acrobat.Pro')
    SystemEvents = app(id='com.apple.systemevents')
    acro_process = SystemEvents.application_processes[u'Acrobat']

    filename = datetime.now().strftime('Scanned Document %y%m%d %H%M%S')

    SA = ScriptingAddition()
    SA.activate()
    while True:
        result = SA.display_dialog('How many pages do you wish to scan?',
                                   buttons=['Cancel', 'Scan'],
                                   cancel_button=1, default_button=2,
                                   default_answer='1')
        if result is None:
            return False
        try:
            pages = int(result[k.text_returned])
        except ValueError:
            continue
        if pages > 0:
            break

    Acrobat.activate()

    acro_process.menu_bars[1].menu_bar_items['Document'].menus[1].\
        menu_items['Scan to PDF...'].click()
    acro_process.windows['Acrobat Scan'].buttons['Scan'].click()

    # pause (Carbon -> Cocoa? use keystrokes instead?)
    acro_process.windows['Save Scanned File As'].text_fields[1].value.\
        set(filename)
    acro_process.windows['Save Scanned File As'].buttons['Save'].click()

    acro_scan_window = acro_process.windows['Acrobat Scan']

    while True:
        acro_process.windows['DSmobile 600'].buttons['Scan'].click()
        while not acro_scan_window.exists():
            time.sleep(0.1)

        pages -= 1

        if pages == 0:
            acro_scan_window.groups[1].radio_buttons[2].click()
            acro_scan_window.buttons['OK'].click()
            break

        acro_scan_window.groups[1].radio_buttons[1].click()
        acro_scan_window.buttons['OK'].click()

    scanned_document = Acrobat.documents['%s.pdf' % filename]
    scanned_file = scanned_document.file_alias(timeout=0)
    scanned_document.close()

    record = Paper.import_(files=[scanned_file], deleting_afterwards=True)[0]
    contents = record.contents()
    m = re.search('(%s)' % '|'.join(map(re.escape, sources)), contents,
                  re.IGNORECASE)
    if m:
        # use the saved source's case
        title = sources[map(str.lower, sources).index(m.group(1).lower())]
    else:
        title = '???'

    extracted, no_match = extract_date(contents)
    if extracted:
        title += extracted.strftime(' %Y-%m')
        record.creation_date.set(extracted)

    record.title.set(title)

    return True

# update_all()

# XXX incremental source recording from EagleFiler (use tag to record)

while scan_one():
    writePlist({'Sources': sources}, PREFERENCES_PATH)
