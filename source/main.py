#!/usr/bin/env python
#
# The following code is Google App Engine compatible Google Doc Exporter
# for CloudSpokes(TM) Challenge. A live version is available at:
#  http://docexporter.appspot.com/


import urllib2
import unicodedata
import re
from StringIO import StringIO
import zipfile
from xml.dom import minidom

from google.appengine.ext import webapp
from google.appengine.ext.webapp import util
from google.appengine.api import users
from google.appengine.api import urlfetch
from google.appengine.runtime import DeadlineExceededError

# Page templates (we avoid using django since we essentially have a single frontend page).
PAGE_HEAD = \
"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "DTD/xhtml1-transitional.dtd">
<html>
<head>
<style type="text/css">
body {
    font-family: Arial,Sans-serif;
    font-size: 13px;
}
table.doclist {
    width: 80%;
}
table.doclist th {
    background-color: #eee;
}
table.doclist td,th {
    border-bottom: 1px solid #ddd;
    padding:2px;
    text-align: left;
}
input.dnld {
    margin:5px;
}
table.doclist a {
    text-decoration: none;
    color: black;
    border-bottom: 1px dotted black;
}
table.doclist a:hover {
    border-bottom: 1px solid black;
}
</style>
<script type="text/javascript">
function hl(elem) {
    var bgcolor;
    if (elem.checked) {
        bgcolor = '#f2f2f2';
    } else {
        bgcolor = '#fff';
    }
    var cnodes = elem.parentNode.parentNode.childNodes;
    for(i=0;i<cnodes.length;i++) {
        cnodes[i].style.backgroundColor=bgcolor;
    }
}
function sa(elem) {
    var keynodes = document.getElementsByName("key");
    for (j=0;j<keynodes.length;j++) {
        keynodes[j].checked = elem.checked;
        hl(keynodes[j]);
    }
}
</script>
<title>CloudSpokes Challenge - Google Docs Exporter</title>
</head>
<body>
<h3>Google Doc Exporter</h3>
<p>If you encounter error downloading multiple files, try selecting lesser files or download single files.</p>
"""

PAGE_TAIL = \
"""</body>
</html>
"""

FORM_HEAD = \
"""<form action="/dnldmulti/" method="post">
<input class="dnld" type="submit" value="Download selected">
<table class="doclist" cellspacing="0">
<tr>
<th><input type="checkbox" onchange="sa(this)"></th>
<th>Title</th>
<th>Author Name</th>
<th>Author Email</th>
<th>Type</th>
</tr>
"""

FORM_TAIL = \
"""</table><input class="dnld" type="submit" value="Download selected"></form>
"""

AUTH_ERROR = \
"""<p>Autorization error. Operation failed.</p>
"""

RELOGIN_MESSAGE = \
"""<p>Please <a href="/">relogin</a> to continue.</p>
"""

GENERAL_ERROR = \
"""<p>An error occured. Operation failed.</p>
"""

INVALID_URL = \
"""<p>Invalid data in URL. Operation failed.</p>
"""

LIST_PAGE_LINK = \
"""<p><a href="/list/">Click here</a> to view your list of documents.</p>
"""

SELECT_FILE_MESSAGE = \
"""<p>Please select atleast one document to download.</p>
"""

TIMED_OUT_MESSAGE = \
"""<p>Downloading has taken longer than allowed. Please try again or select another file. If you were downloading multiple files, select lesser files and try again.</p>
"""

# Some config values for different types of documents
doc_config = {
    'document': {
        'url':'https://docs.google.com/feeds/download/documents/Export?docID=%(doc_key)s&exportFormat=doc',
        'content-type':'application/msword',
        'ext':'doc'
    },
    'presentation': {
        'url':'https://docs.google.com/feeds/download/presentations/Export?docID=%(doc_key)s&exportFormat=pdf',
        'content-type':'application/pdf',
        'ext':'pdf'
    },
    'spreadsheet': {
        'url':'https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=%(doc_key)s&exportFormat=xls',
        'content-type':'application/vnd.ms-excel',
        'ext':'xls'
    }
}

def getUrl(url, sessToken):
    return urlfetch.fetch(url=url, headers= {
                'Content-Type':'application/x-www-form-urlencoded',
                'Authorization':'AuthSub token="'+sessToken+'"',
                'GData-Version':'2.0',
                }, deadline=10)

def getTextInTag(elem, tag):
    return elem.getElementsByTagName(tag)[0].childNodes[0].data

def slugify(value):
    value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore')
    value = unicode(re.sub('[^\w\s-]', '', value).strip().lower())
    value = re.sub('[-\s]+','_',value)
    return value

def token_error(resp):
    resp.out.write(PAGE_HEAD+AUTH_ERROR+RELOGIN_MESSAGE+PAGE_TAIL)

def result_error(resp):
    resp.out.write(PAGE_HEAD+GENERAL_ERROR+RELOGIN_MESSAGE+PAGE_TAIL)

def xml_error(resp):
    result_errror(resp)

def doc_data_error(resp):
    resp.out.write(PAGE_HEAD+INVALID_URL+LIST_PAGE_LINK+PAGE_TAIL)

def doc_checkbox_error(resp):
    resp.out.write(PAGE_HEAD+SELECT_FILE_MESSAGE+LIST_PAGE_LINK+PAGE_TAIL)

def deadline_exceeded_error(resp):
    resp.out.write(PAGE_HEAD+TIMED_OUT_MESSAGE+LIST_PAGE_LINK+PAGE_TAIL)

class MainHandler(webapp.RequestHandler):
    def get(self):
        user = users.get_current_user()
        if not user:
            self.redirect(users.create_login_url(self.request.uri))
        else:
            self.redirect("https://www.google.com/accounts/AuthSubRequest?scope=https%3A%2F%2Fdocs.google.com%2Ffeeds%2F%20https%3A%2F%2Fspreadsheets.google.com%2Ffeeds%2F&session=1&secure=0&next=http%3A%2F%2Fdocexporter.appspot.com%2Fgetsesstoken%2F")
            # Use the following line when testing in localhost
            #self.redirect("https://www.google.com/accounts/AuthSubRequest?scope=https%3A%2F%2Fdocs.google.com%2Ffeeds%2F%20https%3A%2F%2Fspreadsheets.google.com%2Ffeeds%2F&session=1&secure=0&next=http%3A%2F%2Flocalhost:8080%2Fgetsesstoken%2F")

class GetSessTokenHandler(webapp.RequestHandler):
    def get(self):
        singleUseToken = self.request.get("token")
        url = "https://www.google.com/accounts/AuthSubSessionToken"
        result = urlfetch.fetch(url=url, headers= {
                                        'Content-Type':'application/x-www-form-urlencoded',
                                        'Authorization':'AuthSub token="'+singleUseToken+'"',
                                        'Connection':'keep-alive',
                })
        if result.content.startswith("Token="):
            sessToken = result.content[6:]
            self.response.headers.add_header('Set-Cookie','sess_token=%s; path=/'%sessToken)
            self.redirect('/list/')
        else:
            token_error(self.response)

class ListHandler(webapp.RequestHandler):
    def get(self):
        sessToken = self.request.cookies.get('sess_token')
        url = "https://docs.google.com/feeds/documents/private/full"
        result = getUrl(url, sessToken)
        if result.status_code != 200:
            result_error(self.response)
        XML_SIGNATURE = "<?xml version='1.0' encoding='UTF-8'?>"
        if result.content.startswith(XML_SIGNATURE):
            doc_list_xml = result.content[len(XML_SIGNATURE):]
            doc_list_parsed = minidom.parseString(doc_list_xml)
            entry_list = doc_list_parsed.documentElement.getElementsByTagName("entry")
            self.response.out.write(PAGE_HEAD+FORM_HEAD)
            for entry in entry_list:
                doc_id = entry.getElementsByTagName("id")[0].childNodes[0].data
                doc_type,doc_key = doc_id[doc_id.rfind("/")+1:].split("%3A")
                if doc_type not in doc_config.keys():
                    continue
                title = getTextInTag(entry, "title")
                author = entry.getElementsByTagName("author")[0]
                name = getTextInTag(author, "name")
                email = getTextInTag(author, "email")
                doc_name = slugify(title)+"."+doc_config[doc_type]['ext']
                self.response.out.write(\
'<tr>'+\
'<td><input type="checkbox" onchange="hl(this);" name="key" value="'+doc_type+'%3A'+doc_key+'%3A'+doc_name+'">'+\
    '<img src="/static/'+doc_type+'.png"></td>'+\
'<td><a href="/download/?type='+doc_type+'&key='+doc_key+'&name='+doc_name+'">'+title+'</a></td>'+\
'<td>'+name+'</td>'+\
'<td>'+email+'</td>'+\
'<td>'+doc_type+'</td>'+\
'</tr>\n'
                    )
            self.response.out.write(FORM_TAIL+PAGE_TAIL)
        else:
            xml_error(self.response)

class DownloadHandler(webapp.RequestHandler):
    def get(self):
        try:
            sessToken = self.request.cookies.get('sess_token')
            if (sessToken == None):
                token_error(self.response)
                return
            doc_type = self.request.get('type',None)
            doc_key = self.request.get('key',None)
            doc_name = self.request.get('name',None)
            if (doc_type not in doc_config.keys() or doc_key is None or doc_name is None):
                doc_data_error(self.response)
                return
            doc_url = doc_config[doc_type]['url'] % {'doc_key':doc_key}
            result = getUrl(doc_url, sessToken)
            if result.status_code != 200:
                result_error(self.response)
                return
            self.response.headers['Content-Disposition'] = 'attachment; filename=%s' % doc_name
            self.response.headers['Content-Type'] = doc_config[doc_type]['content-type']
            self.response.out.write(result.content)
        except DeadlineExceededError,e:
            deadline_exceeded_error(self.response)
            return
      
class DownloadMultiHandler(webapp.RequestHandler):
    def post(self):
        try:
            sessToken = self.request.cookies.get('sess_token', None)
            if (sessToken == None):
                token_error(self.response)
                return
            atleast_one = False
            zipstream = StringIO()
            zipfl = zipfile.ZipFile(zipstream,"w")
            files_written = {}
            for key in self.request.POST.getall('key'):
                atleast_one = True
                try:
                    doc_type, doc_key, doc_name = key.split("%3A")
                except ValueError:
                    doc_data_error(self.response)
                    return
                doc_name = str(doc_name)
                if files_written.has_key(doc_name):
                    files_written[doc_name] += 1
                    doc_name = str(files_written[doc_name])+"_"+doc_name
                else:
                    files_written[doc_name] = 1
                doc_url = doc_config[doc_type]['url'] % {'doc_key': doc_key}
                result = getUrl(doc_url, sessToken)
                if result.status_code != 200:
                    result_error(self.response)
                    return
                content =  StringIO(result.content)
                content.seek(0)
                zipfl.writestr(doc_name, content.getvalue())
                
            zipfl.close()
            zipstream.seek(0)
            if not atleast_one:
                doc_checkbox_error(self.response)
                return
            self.response.headers['Content-Type'] = 'application/zip'
            self.response.headers['Content-Disposition'] = 'attachment; filename="documents.zip"'
            self.response.out.write(zipstream.getvalue())
        except DeadlineExceededError,e:
            deadline_exceeded_error(self.response)
            return

def main():
    application = webapp.WSGIApplication([('/', MainHandler),
                                          ('/getsesstoken/', GetSessTokenHandler),
                                          ('/list/', ListHandler),
                                          ('/download/', DownloadHandler),
                                          ('/dnldmulti/', DownloadMultiHandler)],
                                         debug=True)
    util.run_wsgi_app(application)


if __name__ == '__main__':
    main()
