#!/usr/bin/python

from appscript import *
from datetime import datetime
from itertools import izip
from osax import *
from plistlib import readPlist, readPlistFromString, writePlist
from subprocess import call, check_output, CalledProcessError
import aem
import os
import re

PREFERENCES_PATH = \
    os.path.expanduser('~/Library/Preferences/net.sabi.UpdateDates.plist')

DATE_FORMATS = (('%m-%d-%y', r'\d{1,2}-\d{1,2}-\d{2}'           ), # Busey new
                ('%m/%d/%y', r'\d{1,2}/\d{1,2}/\d{1,2}'         ), # T-Mobile
                ('%m.%d.%y', r'\d{1,2}\.\d{1,2}\.\d{1,2}'       ), # iFixit
                ('%b%d,%Y',  r'[A-Z][a-z][a-z] ?\d{1,2}, ?\d{4}'), # AmerenIP
                ('%B%d,%Y',  r'[A-Z][a-z]+ *\d{1,2}, *\d{4}'    ), # Amazon
                ('%B%d.%Y',  r'[A-Z][a-z]+ *\d{1,2}\. *\d{4}'   ), # Amazon
                ('%b%d.%Y',  r'[A-Z][a-z]+ *\d{1,2}\. *\d{4}'   ), # Bloomie's
                ('of%Y%m%d', r'of \d{8}'                        ), # Amazon
                ('%m/%d/%Y', r'\d{1,2}/\d{1,2}/\d{4}'           ), # Busey
                ('%b%d%Y',   r'[A-Z]{3} \d{1,2} \d{4}'          ), # State Farm
                ('%b%d,%Y',  r'[A-Z]{3} \d{1,2}, \d{4}'         ), # State Farm
                ('%d%b%Y',   r'\d{1,2} ?[A-Z][A-Za-z]{2} ?\d{4}'), # Apple
                ('%Y-%m-%d', r'\d{4}-\d{2}-\d{2}'               ), # MacSpeech
                ('%d%B%Y',   r'\d{1,2} *[A-Z][a-z]+ *\d{4}'     ), # Vagabond Inn
                ('%Y-%m',    r'\d{4}-\d{2}'                     ), # title
                # bad OCR formats - keep at bottom
                ('%m1%d/%y', r'\d{1,2}1\d{1,2}/\d{1,2}'         ), # T-Mo bad
                ('%m/%d1%y', r'\d{1,2}/\d{1,2}1\d{1,2}'         ), # T-Mo bad
                ('%m/%d/%y', r'\d{1,2}/ \d{1,2}/ \d{1,2}'       ), # T-Mo bad
                ('%m1%d/%Y', r'\d{2}1\d{2}/\d{4}'               ), # Temple bad
                ('%m/%d1%Y', r'\d{1,2}/\d{1,2}1\d{4}'           ), # TotalVac bad
                ('%m/%d/%Y',
                 r'(?:\d ?){1,2}/ (?:\d ?){1,2}/ (?:\d ?){4}'   ), # Busey bad
                )

TITLE_DATE_FORMATS = DATE_FORMATS + (('%Y', r'20\d{2}'),) # title only

def date_re(formats):
    return re.compile('|'.join(r'(\b%s\b)' % regex
                               for format, regex in formats))

def date_extractor(formats):
    return lambda text, match=None: extract_date(text, match,
                                                 date_re(formats), formats)

extract_date_from_contents = date_extractor(DATE_FORMATS)
extract_date_from_title = date_extractor(TITLE_DATE_FORMATS)

def extract_date(text, match, re_date, formats):
    no_format = []
    for m in re_date.finditer(text):
        matched_format = m.lastindex
        format = formats[matched_format - 1][0]
        matched = m.group(matched_format).replace(' ', '')
        try:
            parsed = datetime.strptime(matched, format)
        except ValueError as e: # not a date
            no_format.append((matched, format, e))
            continue
        if not match or (match.year, match.month) == (parsed.year, parsed.month):
            if 1990 < parsed.year < 2100:
                return parsed.date(), no_format
        no_format.append(m.group(matched_format))
    return None, no_format

RE_TITLE_DATE = date_re(TITLE_DATE_FORMATS)

def extract_source_from_title(title, title_date):
    if title_date:
        return title[:RE_TITLE_DATE.search(title).start(0)].rstrip()
    else:
        return title

EagleFiler = app(id='com.c-command.EagleFiler')
Paper = EagleFiler.library_documents['Paper.eflibrary']

def read_sources():
    return map(unicode, readPlist(PREFERENCES_PATH).get('Sources', []))

def write_sources():
     writePlist({'Sources': sources}, PREFERENCES_PATH)

def add_source(source, contents):
    source = unicode(source)
    if source and re.search(re.escape(source), contents, re.IGNORECASE):
        if source in sources:
            sources.remove(source)
        else:
            return True # source is new
        sources.insert(0, source)  # most recently referenced ones at top

def has_encoding_application(path, encoding_application):
    try:
        metadata = readPlistFromString(check_output(['/usr/bin/mdls', '-plist', '-', path]))
    except CalledProcessError:
        return False
    if not isinstance(metadata, dict):
        return False
    return encoding_application in metadata.get('kMDItemEncodingApplications', [])

def update_all():
    record_count = 0
    no_regex_count = 0
    no_format_count = 0
    impossible_count = 0
    new_sources = []

    record_ids = Paper.library_records.id()
    record_utis = Paper.library_records.universal_type_identifier()

    for record_id, record_uti in izip(record_ids, record_utis):
        if record_uti != 'com.adobe.pdf':
            continue

        record = Paper.library_records.ID(record_id)
        tags = record.assigned_tag_names()
        if 'impossible' in tags:
            continue # OCR inadequate/data missing from document

        record_count += 1
        title = record.title()
        title_date, no_format = extract_date_from_title(title)
        source = extract_source_from_title(title, title_date)

        contents = record.text_content()

        if add_source(source, contents):
            new_sources.append(source)

        contents_date, no_format = extract_date_from_contents(contents,
                                                              title_date)

        if not contents_date:
            print '%s (extracted: %s)' % (title, title_date)
            for nf in no_format:
                print '  ', nf
            if not title_date:
                continue

            if no_format:
                no_format_count += 1
                tags.append('no_format')
            else:
                no_regex_count += 1
                tags.append('no_regex')

            record.note_text.set(contents)
            Paper_window.selected_records.set([record])
            EagleFiler.activate()
            record.assigned_tag_names.set(tags)

            disposition = raw_input()
            if disposition == 'i':
                tags.append('impossible')
                record.note_text.set('')
                record.assigned_tag_names.set(tags)
            elif disposition == 'd':
                while True:
                    date_format = raw_input('date format: ')
                    if not date_format: break
                    regex = raw_input('regex: ')
                    if not regex: break
                    date_formats = ((date_format.replace(' ', ''), regex),)
                    print extract_date(contents, title_date,
                                       re_wrap(date_formats), date_formats)
            elif disposition == 'q':
                return

        record.creation_date.set(contents_date or title_date)
        # print 'date:', contents_date or title_date

    print
    print '-' * 50
    print '   Total records:', record_count
    print '  No regex match:', no_regex_count
    print ' No format match:', no_format_count
    print 'Successful match:', record_count - no_regex_count - no_format_count

    print '%d new sources:' % len(new_sources)
    for source in sorted(new_sources):
        print '\t%s' % source

    write_sources()

def title_date_record(record):
    Paper_window.selected_records.set([record])

    title = record.title()
    contents = record.text_content()
    date, no_format = extract_date_from_contents(contents)
    title_date, no_format = extract_date_from_title(title)

    if not title_date:
        m = re.search('(%s)' % '|'.join(map(re.escape, sources)), contents,
                      re.IGNORECASE)
        if m:
            # use the saved source's case
            title = sources[map(unicode.lower, sources).index(m.group(1).lower())]
        else:
            title = '???'

        if date:
            title += date.strftime(' %Y-%m')

    EagleFiler.activate()

    SA = ScriptingAddition()
    SA.activate()
    result = SA.display_dialog('Title this document:',
                               buttons=['Cancel', 'Title'],
                               cancel_button=1, default_button=2,
                               default_answer=title)
    if result is None:
        return

    title = result[k.text_returned]
    title_date, no_format = extract_date_from_title(title)

    if title_date and (not date or (title_date.year, title_date.month) !=
                                   (date.year, date.month)):
        date = title_date

    if date:
        record.creation_date.set(date)

    if add_source(extract_source_from_title(title, title_date),
                  record.text_content()):
        write_sources()

    record.title.set(title)
    record.filename.set(title)

def optimize_record(record):
    Acrobat = app(id='com.adobe.Acrobat.Pro')
    SystemEvents = app(id='com.apple.systemevents')
    acro_process = SystemEvents.application_processes[u'Acrobat']

    file = record.file()
    filename = os.path.basename(file.path)
    creator = SystemEvents.files[file.hfspath].creator_type()

    if creator == 'CARO':
        return # already written by Acrobat

    if not has_encoding_application(file.path, 'IJ Scan Utility'):
        return # not a scanned document

    Acrobat.activate()
    Acrobat.open(record.file())

    acro_process.menu_bars[1].menu_bar_items['Document'].menus[1].\
        menu_items['Optimize Scanned PDF'].click()
    acro_process.windows['Optimize Scanned PDF'].buttons['OK'].click()

    Acrobat.documents[filename].save(to=file)
    Acrobat.documents[filename].close()

def update_selected():
    selected_records = Paper_window.selected_records()

    for record in selected_records:
        title_date_record(record)

    for record in selected_records:
        if record.universal_type_identifier() != 'com.adobe.pdf':
            continue
        optimize_record(record)

if __name__ == '__main__':

    if not Paper.exists():
        EagleFiler.open(os.path.expanduser('~/Documents/Paper/Paper.eflibrary'))

    # XXX filtering doesn't work, even in AppleScript
    # Paper_window = EagleFiler.browser_windows[aem.its.property(aem.app.elements('docu')).eq(Paper)]()

    # appscript gets confused between the property and class 'document'
    window_documents = EagleFiler.AS_newreference(
        aem.app.elements('BroW').property('docu'))()
    for window_index, window_document in enumerate(window_documents):
        if window_document == Paper:
            # we can't store a persistent reference, because the class returned
            # is '\0\0\0\0'
            # Paper_window = EagleFiler.browser_windows[window_index + 1].get()
            Paper_window = EagleFiler.browser_windows.ID(
                EagleFiler.browser_windows[window_index + 1].id.get())

    if os.path.exists(PREFERENCES_PATH):
        try:
            sources = read_sources()
        except:
            call(['/usr/bin/plutil', '-convert', 'xml1', PREFERENCES_PATH])
            sources = read_sources()
    else:
        sources = []

    # update_all()
    update_selected()

    EagleFiler.activate()

# XXX incremental source recording from EagleFiler (use tag to record)
