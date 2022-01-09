from html.parser import HTMLParser
from email.header import Header, decode_header
import mailbox
import base64
import quopri
import re
import sys
import html2text
import bs4
import csv 

""" ____Format utils____ """

class MLStripper(HTMLParser):
    """
    Strip HTML from strings in Python
    https://stackoverflow.com/questions/753052/strip-html-from-strings-in-python
    """
    def __init__(self):
        super().__init__()
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)



def strip_tags(html):
    """
    Use MLStripper class to strip HMTL from string
    """
    s = MLStripper()
    s.feed(html)
    return s.get_data()


def strip_payload(payload):
    """
    Remove carriage returns and new lines
    """
    return payload.replace('\r', ' ').replace('\n', ' ')


def encoded_words_to_text(encoded_words):
    """
    Not used, left for reference only
    https://dmorgan.info/posts/encoded-word-syntax/
    """
    encoded_word_regex = r'=\?{1}(.+)\?{1}([B|Q])\?{1}(.+)\?{1}='
    # encoded_word_regex = r'=\?{1}.+\?{1}[B|Q|b|q]\?{1}.+\?{1}='
    charset, encoding, encoded_text = re.match(encoded_word_regex, encoded_words, re.IGNORECASE).groups()
    if encoding.upper() == 'B':
        byte_string = base64.b64decode(encoded_text)
    elif encoding.upper() == 'Q':
        byte_string = quopri.decodestring(encoded_text)
    return byte_string.decode(charset)



""" ____Custom Message class____ """

class CustomMessage():
    """
    The CusomMessage class represents an email message with three fields:
    - :body:
    - :subject:
    - :content_type: (document, plain text, HTML, image...)
    """
    def __init__(self, body, subject, content_type):
        """
        Constructor
        It tries to find the subject's encoding and decode it accordingly
        It decodes the body based on the content type
        """
        self.content_type = content_type

        # Decode subject if encoded in utf-8
        if isinstance(subject, Header):
            subject = decode_header(subject)[0][0].decode('utf-8')

        # The subject can have several parts encoded in different formats
        # These parts are flagged with strings like '=?UTF-8?'
        if subject is not None and ('=?ISO-' in subject.upper() or '=?UTF-8?' in subject.upper()):
            self.subject = ''
            for subject_part in decode_header(subject):
                # Decode each part based on its encoding
                # The encoding could be returnd by the "decode_header" function
                if subject_part[1] is None:
                    self.subject += strip_payload(subject_part[0].decode())
                else:
                    self.subject += strip_payload(subject_part[0].decode(subject_part[1]))
        elif subject is None:
            # Empty subject
            self.subject = ''
        else:
            # Subject is not encoded or other corner cases that are not considered
            self.subject = strip_payload(subject)

        # Body decoding
        if 'text' in self.content_type:
            # Decode text messages
            try:
                decoded_body = body.decode('utf-8')
            except UnicodeDecodeError:
                decoded_body = body.decode('latin-1')

            if 'html' in self.content_type:
                # If it is an HTML message, remove HTML tags
                h = html2text.HTML2Text()

                h.ignore_links = True
                h.ignore_tables = True
                h.ignore_images = True
                h.ignore_anchors = True
                h.ignore_emphasis = True

                self.body = strip_payload(h.handle(decoded_body))
            else:
                self.body = strip_payload(decoded_body)
        else:
            # If not text, return the body as it is
            self.body = body

    def __str__(self):
        body_length = 2000
        printed_body = self.body[:body_length]
        if 'text' in self.content_type:
            # Shorten long message bodies
            if len(self.body) > body_length:
                printed_body += "..."
        return " ---- Custom Message ---- \n  -- Content Type: {}\n  -- Subject: {}\n  -- Body --\n{}\n\n".format(self.content_type, self.subject, printed_body)

    def get_body(self):
        return self.body

    def get_subject(self):
        return self.subject

    def get_content_type(self):
        return self.content_type

    def create_vector_line(self, label):
        """
        Creates a CSV line with the message's body and given :label:
        Removes any commas from body and label
        """
        return '{body},{label}'.format(body=self.body.replace(',', ''), label=label)

    @staticmethod
    def extract_types_from_messages(messages):
        """
        Takes a list of CustomMessage and extracts all the existing values for content_type
        ['application/ics', 'application/octet-stream', 'application/pdf', 'image/gif', 'image/jpeg',
        'image/png', 'text/calendar', 'text/html', 'text/plain', 'text/x-amp-html']
        """
        types = set()
        for m in messages:
            types.add(m.get_content_type())
        return sorted(types)



""" ____Extraction utils____ """

def extract_message_payload(mes, parent_subject=None):
    """
    Extracts recursively the payload of the messages contained in :mes:
    When a message is embedded in another, it uses the parameter :parent_subject:
    to set the subject properly (it uses the parent's subject)
    """
    extracted_messages = []
    if mes.is_multipart():
        if parent_subject is None:
            subject_for_child = mes.get('Subject')
        else:
            subject_for_child = parent_subject
        for part in mes.get_payload():
            extracted_messages.extend(extract_message_payload(part, subject_for_child))
    else:
        extracted_messages.append(CustomMessage(mes.get_payload(decode=True), parent_subject,  mes.get_content_type()))
    return extracted_messages

def extract_message_payload2(mes, parent_subject=None):
    """
    Extracts recursively the payload of the messages contained in :mes:
    When a message is embedded in another, it uses the parameter :parent_subject:
    to set the subject properly (it uses the parent's subject)
    """
    extracted_messages = []
    if mes.is_multipart():
        if parent_subject is None:
            subject_for_child = mes.get('Subject')
        else:
            subject_for_child = parent_subject
        for part in mes.get_payload():
            extracted_messages.extend(extract_message_payload(part, subject_for_child))
    else:
        extracted_messages.append(CustomMessage(mes.get_payload(decode=True), parent_subject,  mes.get_content_type()))
    return extracted_messages

def text_messages_to_string(mes):
    """
    Returns the email's body extracted from :mes: as a string.
    Ignores images and documents.
    :mes: should be a list of CustomMessage objects.
    """
    output = ''
    for m in mes:
        if m.get_content_type().startswith('text'):
            output += str(m)
    return output


def create_classification_line(mes, label):
    """
    Creates CSV line(s) with two columns: the email's body extracted from :mes:
    and its classification (:label:)
    Ignores images, documents and calendar messages.
    :mes: should be a list of CustomMessage objects.
    """
    output = ''
    for m in mes:
        if m.get_content_type().startswith('text') and m.get_content_type() != 'text/calendar':
            output += m.create_vector_line(label) + '\n'
    return output


def to_file(text, file):
    """
    Writes :text: to :file:
    """
    f = open(file, 'w', encoding='utf-8')
    f.write(text)
    

    f.close


def extract_mbox_file(file):
    """
    Extracts all the messages included in an mbox :file:
    by calling extract_message_payload
    """
    mbox = mailbox.mbox(file)
    messages = []
    for message in mbox:
        messages.extend(extract_message_payload(message))
    return messages



if __name__ == '__main__':
    argv = sys.argv
    if len(argv) != 2:
        print('Invalid arguments')
    else:
        file = argv[1]
        messages = extract_mbox_file(file)


# Call to create a CSV file with the extracted data (body + label)
# to_file(create_classification_line(messages, 'pockets'), file + '_features.csv')
# Call to export all the extracted data
# to_file(text_messages_to_string(messages), file + '_full_extract')

mydictlist = []

def getcharsets(msg):
    charsets = set({})
    for c in msg.get_charsets():
        if c is not None:
            charsets.update([c])
    return charsets

def getBody(msg):
    while msg.is_multipart():
        msg=msg.get_payload()[0]
    t=msg.get_payload(decode=True)
    for charset in getcharsets(msg):
        t=t.decode(charset)
    return t

def get_html_text(html):
    try:
        return bs4.BeautifulSoup(html, 'lxml').body.get_text(' ', strip=True)
    except AttributeError: # message contents empty
        return None

class GmailMboxMessage():
    def __init__(self, email_data):
        if not isinstance(email_data, mailbox.mboxMessage):
            raise TypeError('Variable must be type mailbox.mboxMessage')
        self.email_data = email_data

    def parse_email(self):
        email_labels = self.email_data['X-Gmail-Labels']
        email_date = self.email_data['Date']
        email_from = self.email_data['From']
        email_to = self.email_data['To']
        email_subject = self.email_data['Subject']
        # email_text = self.read_email_payload() 
        email_text = extract_message_payload(self.email_data)[0].body.replace(',', '')
        mydict = {}
        mydict['from'] = email_from
        mydict['subject'] = email_subject
        mydict['text'] = str(email_text)
        mydict['Label'] = "Promotional"
        mydictlist.append(mydict)


    def read_email_payload(self):
        email_payload = self.email_data.get_payload()
        if self.email_data.is_multipart():
            email_messages = list(self._get_email_messages(email_payload))
        else:
            email_messages = [email_payload]
        return [self._read_email_text(msg) for msg in email_messages]

    def _get_email_messages(self, email_payload):
        for msg in email_payload:
            if isinstance(msg, (list,tuple)):
                for submsg in self._get_email_messages(msg):
                    yield submsg
            elif msg.is_multipart():
                for submsg in self._get_email_messages(msg.get_payload()):
                    yield submsg
            else:
                yield msg

    def _read_email_text(self, msg):
        content_type = 'NA' if isinstance(msg, str) else msg.get_content_type()
        encoding = 'NA' if isinstance(msg, str) else msg.get('Content-Transfer-Encoding', 'NA')
        if 'text/plain' in content_type and 'base64' not in encoding:
            msg_text = msg.get_payload()
        elif 'text/html' in content_type and 'base64' not in encoding:
            msg_text = get_html_text(msg.get_payload())
        elif content_type == 'NA':
            msg_text = get_html_text(msg)
        else:
            msg_text = None
        return (content_type, encoding, msg_text)

######################### End of library, example of use below

mbox_obj = mailbox.mbox('C:/Users/mypc/Downloads/mlassgn/final_promotional.mbox')

num_entries = len(mbox_obj)

for email_obj in mbox_obj:
    email_data = GmailMboxMessage(email_obj)
    email_data.parse_email()


for i in range(len(mydictlist)):
    print(mydictlist[i])
# for thisemail in mbox_obj:
#     body = getBody(thisemail)
#     print(body)

  

    
# field names 
fields = ['from', 'subject', 'text', 'Label'] 
    
# data rows of csv file 
    
# name of csv file 
filename = "Promotional.csv"
    
# writing to csv file 
with open(filename, 'w', encoding='utf-8') as csvfile: 
    # creating a csv dict writer object 
    writer = csv.DictWriter(csvfile, fieldnames = fields) 
        
    # writing headers (field names) 
    writer.writeheader() 
        
    # writing data rows 
    writer.writerows(mydictlist) 
