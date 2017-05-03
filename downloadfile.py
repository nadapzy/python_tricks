# -*- coding: utf-8 -*-
"""
Created on Tue May 02 16:26:06 2017

@author: v0d001u
"""

import win32com.client
import os
import shutil
import urllib
#from six.moves import urllib
#from BeautifulSoup import *
from bs4 import BeautifulSoup
import urllib2

os.environ['https_proxy'] = 'proxy1'
os.environ['http_proxy'] = 'proxy1'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case,
                                    # the inbox. You can change that number to reference
                                    # any other folder
global i
messages = inbox.Items
message = messages.GetFirst ()
while message:
    if message.Subject == 'EXT: Your Data Insight job is ready: Bi-Weekly In-Pipeline Report':
        attachment = message.attachments
        attachment = attachment.Item(1)
        attachment.SaveASFile(os.getcwd() + '\\' + str(attachment))
#        print (attachment)
    message = messages.GetNext ()

proxies = {'https':r'http://userid:pwd@proxy:8080','http':r'http://userid:pwd@proxy:8080'}



#proxy=urllib.request.ProxyHandler(proxies)
#opener=urllib.request.build_opener(proxy)
#urllib.request.install_opener(opener)

def test():
    link='http://cpanratings.perl.org/csv/all_ratings.csv'
    #urllib.urlretrieve(link,filename=file_name)
    filehandle=urllib.urlopen(link,proxies=proxies)
    with open('temp.csv','wb') as f:
        f.write(filehandle.read())
#test()
        
def test1():
    link='http://cpanratings.perl.org/csv/all_ratings.csv'
    #urllib.urlretrieve(link,filename=file_name)
    filehandle=urllib.URLopener(proxies=proxies)
    filehandle.retrieve(link,file_name)
#    with open('temp.csv','wb') as f:
#        f.write(filehandle.read())
#test1()        

file_name='temp.csv'
text_files = [f for f in os.listdir(os.getcwd()) if f.endswith('.html')]
for fil in text_files:
    print fil
    fhand = urllib.urlopen(fil).read()
    soup = BeautifulSoup(fhand)
    tags = soup('a')
    for tag in tags:
        link = tag.get('href',None)
        filehandle=urllib.URLopener(proxies=proxies)
        filehandle.retrieve(link,file_name)
	  #opn = urllib.urlopen(link,proxies=proxies).read()
    print link




