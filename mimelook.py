#!/usr/bin/env python3

import os
import sys
import re
import base64
import subprocess
import html
from html.parser import HTMLParser

import mailparser  # mail-parser
import markdown  # Markdown
import magic  # python-magic (for mimetypes)


QUOTE_ESCAPE = "MIMELOOK_QUOTES"
mime = magic.Magic(mime=True)


def export_inline_attachments(message, dstdir):
    # find inline attachments in plaintext version
    #inlines = re.findall("\[cid:.*?\]", message.body.split("--- mail_boundary ---")[0])
    # .. or html version (probably safer?):
    inlines = re.findall("src=\"cid:.*?\"", message.body.split("--- mail_boundary ---")[1])

    # return list of tuples: (attachment id, exported file)
    ret = []

    for inline in inlines:
        # find filename
        name_match = re.search("cid:.*@", inline)

        # find content id
        id_match = re.search("@.*", inline)
        attachment_id = inline[id_match.start()+1:-1]

        # find corresponding attachment in the message
        attachment_name = inline[name_match.start()+4:name_match.end()-1]
        attachment = [x for x in message.attachments if x["content-id"].startswith("<{}".format(attachment_name))]
        assert len(attachment) == 1, "Could not get attachment '{}'".format(attachment_name)
        attachment = attachment[0]

        # base64 decode the file and place it in dstdir
        assert attachment["content_transfer_encoding"] == "base64", "Only base64 currently supported"
        b = base64.decodebytes(bytes(attachment["payload"], "ascii"))
        dstfile = os.path.join(dstdir, attachment_name)
        with open(dstfile, "wb") as f:
            f.write(b)

        # same attachment might occur multiple times - only add it once
        att = (attachment_id, dstfile)
        if att not in ret:
            ret.append(att)

    return ret

# make "from: x, sent: d/m-y, to: z, ..." section in outlook style
# grabbed from what outlook webmail does
def format_insane_outlook_header(fromaddr, sent, to, cc, subject):
    ret = """<hr style="display:inline-block;width:98%" tabindex="-1">
<div id="divRplyFwdMsg" dir="ltr"><font face="Calibri, sans-serif" style="font-size:11pt" color="#000000"><b>From:</b> {}<br>
<b>Sent:</b> {}<br>
""".format(fromaddr, sent)

    if to is not None:
        ret += "<b>To:</b> {}<br>\n".format(to)

    if cc is not None:
        ret += "<b>Cc:</b> {}<br>\n".format(cc)

    ret += """<b>Subject:</b> {}</font>
<div>&nbsp;</div>
</div>
""".format(subject)

    return ret


# get message from id
def message_from_msgid(msgid):
    mucmd = "mu find msgid:{} --fields 'l'".format(msgid)
    p = subprocess.Popen(mucmd.split(" "), stdout=subprocess.PIPE)
    messagefiles, _ = p.communicate()
    messagefiles = messagefiles.decode("utf-8").split("\n")

    # check return code ok
    assert p.returncode == 0, "mu find failed"

    # expecting list of at least 1 message and one empty line
    assert(len(messagefiles) > 1), "mu found no messages"

    # use first hit and strip surrounding '
    messagefile = messagefiles[0][1:-1]

    # parse message and grab HTML
    message = mailparser.parse_from_file(messagefile)

    return message

# create crazy outlook-style html reply from message id and the desired html message
def format_outlook_reply(message, htmltoinsert):
    message_html = message.body.split("--- mail_boundary ---")[1]

    # convert CRLF to LF
    message_html = message_html.replace("\r\n", "\n")

    # grab header info
    message_from = message.headers["From"]
    message_to = message.headers["To"] if "To" in message.headers else None
    message_subject = message.headers["Subject"]
    message_date = message.date.strftime("%d %B %Y %H:%M:%S")
    message_cc = message.headers["CC"] if "CC" in message.headers else None

    outlook_madness = format_insane_outlook_header(message_from, message_date,
                                                   message_to, message_cc, message_subject)

    # find body tag in html
    m = re.search("<body.*?>", message_html)
    assert m is not None, "No body tag found in parent HTML"

    # format resulting html email:
    # ..<body> from email being replied to
    # reply message
    # "from yada yada" section
    # remainder of email being replied to
    html = "{}\n{}\n{}\n{}".format(message_html[:m.end()],
                                   htmltoinsert,
                                   outlook_madness,
                                   message_html[m.end():])

    return html


# Convert "> "-style quotes into something else that passes untouched through html.escape()
# Escaped quotes look like this: [[MIMELOOK_QUOTES|X]] where X denotes the
# number of quotes that have been escaped
def escape_quotes(plaintext):
    ret = ""
    for line in plaintext.split("\n"):
        if line.startswith(">"):
            i = 0
            while i < len(line) and line[i] == ">":
                i += 1
            ret += "[[{}|{}]]".format(QUOTE_ESCAPE, i)
            ret += line[i:] + "\n"
        else:
            ret += line + "\n"

    return ret


# Convert previously escaped "> "-style quotes back to their original form.
def unescape_quotes(string):
    retstr = ""
    i = 0

    while i < len(string):
        # find start position of next escaped quote group
        p = string[i:].find("[[{}|".format(QUOTE_ESCAPE))

        if p < 0:
            # no more escaped quotes - grab the rest of the string and return
            retstr += string[i:]
            break

        # found some escaped quotes, grab content of string upto them
        retstr += string[i:i+p]

        # find end of the escaped quote tag
        tag_end_pos = string[i+p:].find("]]")

        # grab the number from the tag
        nquotes = int(string[i+p+2+len(QUOTE_ESCAPE)+1:i+p+tag_end_pos])

        # append that number of quotes
        retstr += ">"*nquotes

        # advance to the character just after the tag
        i += p+tag_end_pos+2

    return retstr


def escape_signature_linebreaks(plaintext):
    m = re.search("^-- ", plaintext, re.MULTILINE)
    if m is not None:
        content = plaintext[:m.start()]
        signature = plaintext[m.start():]
        signature = signature.replace("\n", "  \n")
        return content + signature
    else:
        return plaintext


# Find MIME parts in the plaintext
# Returns plaintext without parts and list of parts
# Warning: assumes only valid <#part...><#/part> tags after the first occurence
# of "<#part" !
def find_mime_parts(plaintext):
    parts = re.findall("<#part.*?<#/part>", plaintext, re.DOTALL)
    m = re.search("<#part.*?<#/part>", plaintext, re.DOTALL)
    if m is not None:
        text = plaintext[:m.start()]
    else:
        text = plaintext
    return text, parts


# Escape HTML. Don't escape lines that start with four spaces, since
# these will be wrapped by <pre> tags by the Markdown-to-html
# conversion.
def html_escape(text):
    ret = ""
    for line in text.split("\n"):
        if not line.startswith("    "):
            ret += html.escape(line) + "\n"
        else:
            ret += line + "\n"

    return ret


# Take desired plaintext message and id of message being replied to
# and format a multipart message with sane plaintext section and
# insane outlook-style html section. The plaintext message is converted
# to HTML supporting markdown syntax.
def plain2fancy(plaintext, msgid):

    # find and strip MIME parts in the ending of the plaintext
    plaintext, parts = find_mime_parts(plaintext)

    # escape HTML in the plaintext, handling quoted content explicitly
    escaped_plaintext = unescape_quotes(html_escape(escape_quotes(plaintext)))

    # handle signature - we expect linebreaks to be preserved in the signature,
    # but let everything else wrap (reminder: Markdown preserves linebreaks if
    # there's two spaces at the end of a line)
    escaped_plaintext = escape_signature_linebreaks(escaped_plaintext)

    # plaintext is converted to html, supporting markdown syntax
    # loosely inspired by http://webcache.googleusercontent.com/search?q=cache:R1RQkhWqwEgJ:tess.oconnor.cx/2008/01/html-email-composition-in-emacs
    text2html = markdown.markdown(escaped_plaintext)

    # get message from message id
    message = message_from_msgid(msgid)

    # insane outlook-style html reply
    madness = format_outlook_reply(message, text2html)

    # find inline attachments and export them to a temporary dir
    attdir = "/dev/shm/mu4e-{}".format(msgid)
    if not os.path.isdir(attdir):
        os.mkdir(attdir)
    attachments = export_inline_attachments(message, attdir)

    # build string of <#part type=x filename=y disposition=inline><#/part> for each
    # attachment, separated by newlines
    attachment_str = ""
    for attachment in attachments:
        mimetype = mime.from_file(attachment[1])
        attachment_str += "<#part type=\"{}\" filename=\"{}\" disposition=inline id=\"{}@{}\"><#/part>\n"\
            .format(mimetype, attachment[1], os.path.basename(attachment[1]), attachment[0])

    # also include attachments that were already present in the plaintext
    attachment_str += "\n".join(parts)

    # write html message to file for inspection before sending
    with open("/dev/shm/mimelook-madness.html", "w") as f:
        f.write(madness)

    # return the multipart message
    multimsg = """<#multipart type=alternative>
<#part type=text/plain>
{}
<#/part>
<#part type=text/html>
{}
<#/part>
<#/multipart>
{}""".format(plaintext, madness, attachment_str)

    return multimsg

if __name__ == '__main__':
    stdin = sys.stdin.read()
    msgid_end_char = stdin.find("\n")
    # expecting first line to be the message id
    msgid = stdin[:msgid_end_char]
    # expecting rest to be plaintext version of the message
    plaintext = stdin[msgid_end_char+1:]

    print(plain2fancy(plaintext, msgid))
