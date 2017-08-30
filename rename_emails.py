# -*- coding: utf-8 -*-
"""
Created on Thu Aug 03 14:24:00 2017

@author: king.r.paul@gmail.com

adapted from:
    http://www.media-division.com/using-python-to-batch-rename-email-files/
    for use with Python 2.7
"""

import string
import sys
import email
import glob
import os
import time
from string import whitespace
from email.utils import parsedate_tz as parsedate
from email.utils import mktime_tz as mktime

LINUX = 'linux' in sys.platform
WINDOWS = 'win' in sys.platform

EXT = ".eml"

STAMPS = ['RE-', 'Re-', 'RE:', 'Re:', 'Re ', 'FW:', 'Fwd:', 'FW_',
          'Automatic reply:', 'Notification:', 'DO NOT REPLY', 'Emailing:',
          'Out of Office:', 'REMINDER:', ' Reminder:', 'Accepted:', 'Ad:']
QUOTES = ["'", '"', '&quot']
ENCODINGS = ["=?UTF-8?Q?", "=?UTF-8?q?", "=?iso-8859-1?Q?"]
PUNCTUATION = ['#', '!', '|', '/', '*', '...', '=', '+', '?', ';', ' - ']
ILLEGAL = [bad_char for bad_char in '<>:"/\\|?*\n']
DEAD_SPACE = [w for w in whitespace[:-1]] + ['  ']

NEW_FORMAT = "{} [{}] fm {} to {}"
LINE = '-'*60
NO_ADDR = 'No_Addr'
NO_SUBJ = 'No_subj'
MAX_FILES_ALLOWED = 10000

def remove_bracketed(input_string, left, right):
    """ Remove substring between left and right brackets, inlcuding
    the brackets, from a string and return the remaining string.

        Parameter: string which may contain bracketed section
        """
    bracketed = input_string.split(left)[-1].split(right)[0]
    return input_string.replace(left + bracketed + right, '')

def is_valid_address(addr_string):
    """ Check for valid e-mail address; with one @-sign, with one period and
        no whitespace following and returns a boolean.

        Parameter: string which may be an e-mail address
        """

    if addr_string.count('@') != 1:
        return False
    elif '.' not in addr_string.split('@')[-1]:
        return False
    if any([space in addr_string.split('@')[-1] for space in whitespace]):
        return False
    return True

def clean_name(name_string):
    """ Extract valid e-mail address from string, preferably delimited
        by <braces> and returns it as a string.

        Parameter: string which may contain and e-mail address
        """
    if not name_string:
        return NO_ADDR

    if name_string.count('>') > 1:  # simple handling multiple addresses
        name_string = name_string.split('>')[0] + '> etc'

    try:  # inside braces
        addr = name_string.split('<')[1].split('>')[0]
    except IndexError:
        addr = ''

    if not is_valid_address(addr):
        name_string = remove_bracketed(name_string, '<', '>')
        if is_valid_address(name_string):
            addr = name_string
        else:
            addr = NO_ADDR

    for thing in ILLEGAL:
        addr = addr.replace(thing, '')

    return addr[:35].lower()

def clean_subj(subject):
    """ Format subject string for brevity. """
    if not subject or len(subject) < 2:
        subject = NO_SUBJ

    things_to_remove = STAMPS + ENCODINGS + PUNCTUATION + QUOTES + ILLEGAL
    for thing in things_to_remove:
        subject = subject.replace(thing, '')

    subject = remove_bracketed(subject, "(", ")")
    subject = remove_bracketed(subject, "[", "]")

    subject = subject.replace(' ', '_').replace('__', '_').strip('_')
    return subject[:30]

def get_unopenable(path):
    """ Create list of files that cannot be opened. """
    unopenable = []
    for file_path in glob.glob(os.path.join(path, '*' + EXT)):
        try:
            open(file_path).close()
        except IOError:
            try:
                unopenable.append(file_path)
                continue
            except IOError:
                unopenable.append('?')
                continue
    return unopenable

def get_all_from(path):
    """ Create list of unique senders from e-mails.  """
    senders = []
    for file_path in glob.glob(os.path.join(path, '*' + EXT)):
        try:
            with open(file_path) as open_file:
                msg = email.message_from_file(open_file)
                senders.append(clean_name(msg['From']))
        except IOError:
            continue
        senders.append(clean_name(msg['From']))
    senders = set(senders)
    senders = list(senders)
    senders.sort()
    return senders

def get_all_to(path):
    """ Create list of unique recipients from e-mails.  """
    recipients = []
    for file_path in glob.glob(os.path.join(path, '*' + EXT)):
        try:
            with open(file_path) as open_file:
                try:
                    msg = email.message_from_file(open_file)
                    recipients.append(clean_name(msg['To']))
                except IOError:
                    continue
        except IOError:
            continue
    recipients = set(recipients)
    recipients = list(recipients)
    recipients.sort()
    return recipients

def sanitize_filenames(path):
    """ Rename files in path by stripping out problematic characteers,
        using linux. """
    num_renamed = 0
    allowed = string.whitespace + \
          string.letters + \
          string.digits + \
          string.punctuation
    items = os.listdir(path)        
    for filename in items:
        # remove leading period '.'
        if filename[0] == '.':
            try:
                os.rename(os.path.join(path, filename), 
                          os.path.join(path, filename[1:]))
                filename = filename[1:]
            except IOError:
                pass
        # remove chars with windows problems
        if LINUX:
            sanitized = "".join([c for c in filename if c in allowed]).rstrip()
            if filename != sanitized:
                try:
                    os.rename(os.path.join(path, filename), 
                              os.path.join(path, sanitized))
                    num_renamed += 1    
                except IOError:
                    pass
    return ("Names of {} files were sanitized.".format(num_renamed))

def get_folder_summary(path):
    """ Summarize contents of a folder & return the result as as string."""

    isfile = lambda f: os.path.isfile(os.path.join(path, f))
    isdir = lambda f: os.path.isdir(os.path.join(path, f))
    ismail = lambda f: os.path.splitext(f)[-1].lower() == EXT.lower()

    items = os.listdir(path)
    files = [f for f in items if isfile(f)]
    dirs = [f for f in items if isdir(f)]
    emails = [f for f in items if ismail(f)]

    bad_items = [f for f in items if f not in files and f not in dirs]

    result = 'Path -> "{}"\n'.format(path) + \
    '{:#8} items\n'.format(len(items)) + \
    ' '*4 + '-'*11 + '\n' + \
    '{:#8} files\n'.format(len(files)) + \
    '{:#8} emails\n'.format(len(emails)) + \
    '{:#8} dirs\n'.format(len(dirs)) + \
    '{:#8} bad items\n'.format(len(dirs)) + \
    ' '*4 + '-'*11 + '\n\n    ' + \
    '{}\n'.format('\n    '.join(bad_items))
    return result

def rename_emails(path, dryrun = True):
    """ Rename all e-mail files in folder with informative name.

        No input file-name may be blank or '...'. or non-western symols.

        Writes a log file to the path location and prints 'All done!'
        to the terminal when complete.

        Keyword Arguments:
            dryrun - Boolean, simulation if False, files renamed if True
        """
    all_rename_fails = []
    
    with open(os.path.join(path, 'convert_log.txt'), 'w') as log:

        files = glob.glob(os.path.join(path, '*' + EXT))

        if len(files) > MAX_FILES_ALLOWED:
            print 'Too many files in this folder. Exiting.'
            sys.exit()

        all_from = get_all_from(path)
        all_to = get_all_to(path)
        unopenable = get_unopenable(path)

        log.write(
         'FILE RENAME SUMMARY\n' + LINE + '\n' +
         '# e-mail files founds: {}\n'.format(len(files)) +
         '# unopenable files:    {}\n'.format(len(unopenable)) +
         '# unique senders:      {}\n'.format(len(all_from)) +
         '# unique recipients:   {}\n'.format(len(all_to)) +
         '\n\nUnopenable Files\n' + LINE + '\n' + '\n'.join(unopenable) +
         '\nUnique Senders\n' + LINE + '\n' + '\n'.join(all_from) +
         '\n\nUnique Recipients\n' + LINE + '\n' + '\n'.join(all_to) +
         '\n\nINDIVIDUAL FILES RENAMED\n' + LINE )

        for file_path in files:

            try: # parse the email
                with open(file_path) as open_file:
                    msg = email.message_from_file(open_file)
                    log.write('\nconverting from : ' +
                               os.path.basename(file_path)[:65] + '\n')
            except IOError:
                continue

            mail = dict() # get from, to, subj, date
            mail['from'] = clean_name(msg['From'])
            mail['to'] = clean_name(msg['To'])
            mail['subj'] = clean_subj(msg['Subject'])
            try: #mail['date']
                timestamp = mktime(parsedate(msg['Date']))
                timestring = time.gmtime(timestamp)
                mail['date'] = time.strftime("%Y_%m_%d_%H%M%S", timestring)
            except TypeError:
                log.write('* No Date Found *')
                mail['date'] = "0000_00_00"

            # format the new name
            base_name = NEW_FORMAT.format(mail['date'], mail['subj'],
                                          mail['from'], mail['to'])

            for item in DEAD_SPACE:   # remove whitespace
                while item in base_name:
                    base_name = base_name.replace(item, ' ').strip()

            base_name = base_name[:190]  # truncate
            new_file_path = os.path.join(path, base_name + EXT)

            # ignore if already formatted
            if (os.path.basename(file_path) == base_name + EXT):
                log.write('** No Conversion Needed\n')
                continue

            i = 1   # if naming conflict, append a number
            while(os.path.isfile(new_file_path)):
                new_file_path = os.path.join(path,
                                             base_name + " (" + str(i) + ")" +
                                             EXT)
                i = i+1
           
            try: # rename the file
                log.write('converting to: ' + base_name[:60] + EXT + '\n')
                if not dryrun:
                    try: 
                        os.utime(file_path, (timestamp, timestamp))
                    except UnboundLocalError:
                        pass
                    try:                         
                        os.rename(file_path, new_file_path)
                    except OSError:
                        log.write(' Conversion Fail\n')
                        all_rename_fails.append(file_path)
            except:
                raise AssertionError

        log.write('\n# failed to rename: {}\n'.format(len(all_rename_fails)))
        log.write('\n' + LINE + '\n\n')
        log.write('\n'.join(all_rename_fails) + '\n')

    return '{} e-mail files were processed. \n'.format(len(files))

if __name__ == "__main__":
    if LINUX:
        root_path = "//media//oncobot//SEAGATE//Email_Archive/Years - Copy"
    elif WINDOWS:
        root_path = "G:\\Email_Archive\\YEARS - copy"
    else:
        raise AssertionError
 
#    all_paths = [os.path.join(d, x)
#                 for d, dirs, files in os.walk(root_path)
#                 for x in dirs]

    all_paths = [r'C:\Users\Oncobot\Desktop\Emails']

    for  path in all_paths:
        print LINE
        os.chdir(path)
        print sanitize_filenames(path)
        print get_folder_summary(path = path)
        print rename_emails(path, dryrun = False)
    print '\nAll done!\n' + LINE
