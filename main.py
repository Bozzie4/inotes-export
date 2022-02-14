#!/usr/bin/env python
"""
Usage:  main.py
        main.py [--out-dir=TARGETDIR --mailfile=MAILFILE --mailfolder=FOLDER --cookie=COOKIE --debug]

Options:
  --debug   Print debug information
  --mailfile=MAILFILE   URL to your mailfile as reported by iNotes. Starts with https , ends with .nsf
  --mailfolder=FOLDER   Folder to load (defaults to inbox)
  --cookie=COOKIE   Cookie string as taken from Developer Tools
  --out-dir=DIRECTORY   Directory to store the resulting yaml and conf files in (default: temp or tmp)  Eg. to store in current directoy : --out-dir=.
  -h --help     Show this screen.

"""
import requests
import tempfile
import html
import xmltodict
#from http.cookies import SimpleCookie
from docopt import docopt
import os.path

def prepareCookies(cookieString, debug=False):
    _cookies = {}
    for x in cookieString.split("; "):
        y = x.split("=")
        if debug:
            print(f"cookie : {y[0]} - {y[1]}")
        _cookies[y[0]] = y[1]
    return _cookies

def loadFolder(mailfile, cookies, headers, foldername='($Inbox)', max=0, count=500, start=1):
    if foldername == None:
        foldername = '($Inbox)'
    url = f"{mailfile}/iNotes/Proxy/?OpenDocument&Form=s_ReadViewEntries&PresetFields=FolderName;{foldername},UnreadCountInfo;1,SearchSort;DateA,s_UsingHttps;1,hc;$98,noPI;1&TZType=UTC&Count={count}&Start={start}&resortdescending=5"
    print(f"{url}")
    #print(f"{cookies.keys()}")
    response = requests.request("GET", url, cookies=cookies, headers=headers, allow_redirects=False)
    # returns xml
    viewentries = xmltodict.parse(response.text)
    # get the toplevelentries
    numentries = int(viewentries['readviewentries']['viewentries']['@toplevelentries'])
    print(f"Total entries: {numentries}")
    _unids = []
    _index = 0

    for viewentry in viewentries['readviewentries']['viewentries']['viewentry']:
        #print(f"{viewentry['@unid']}")
        _unids.append(viewentry['@unid'])
        _index = _index+1
        #step out if we reach max
        if max > 0 and (_index+start) > max:
            print(f"Hit max value {max}")
            break
    # Recursively load until max is reached
    start = start + _index
    if (max == 0 or start < max) and (start < numentries):
        _unids = _unids + loadFolder(mailfile, cookies, headers, foldername=foldername, max=max, count=count, start=start)
    return _unids

def main():
    args = docopt(__doc__)

    debug = False
    outdir = None
    mailfile = None
    mailfolder = '($Inbox)'
    cookies = {}
    if args['--debug']:
        debug=True
    if args['--out-dir']:
        outdir = args['--out-dir']
    if args['--mailfile']:
        mailfile = args['--mailfile']
    if args['--mailfolder']:
        mailfolder = args['--mailfolder']
    if args['--cookie']:
        cookieInput = args['--cookie']
        cookies = prepareCookies(cookieInput, debug)
    if mailfile is None or cookies == {}:
        print("--mailfile and --cookie are mandatory options!")
        exit(9)

    payload={}
    headers = {
     'Content-Type': 'application/json',
     'User-Agent': 'Mozilla/5.0 (X11; Fedora; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.80 Safari/537.36'
    }

    if outdir == None:
        outdir = tempfile.gettempdir()
    if outdir.endswith('/'):
        outdir = outdir[:-1]

    notesunids = loadFolder(mailfile, cookies, headers, foldername=mailfolder, max=0)

    print(f"Number of unids loaded {len(notesunids)}\n")

    if len(notesunids) > 0:
        _t = 0
        for unid in notesunids:
            outfilename = f"{outdir}/{unid}.eml"
            _t = _t+1
            if os.path.exists(outfilename):
                if debug:
                    print(f"Skip import - {outfilename} exists . ")
            else:
                url = f"{mailfile}/($All)/{unid}/?OpenDocument&Form=l_MailMessageHeader&PresetFields=FullMessage;1"
                response = requests.request("GET", url, cookies=cookies, headers=headers, data=payload, allow_redirects=False)
                #print(response.text)
                _mail = response.text.replace("<br>","")
                _mail = html.unescape(_mail)
                # DO NOT WRITE EMPTY FILE
                if _mail is None or _mail.strip() == "":
                    print(f"CANNOT PROCESS empty mime : {mailfile}/($All)/{unid}/?OpenDocument")
                else:
                    outf = open(outfilename, "w", encoding='utf-8')
                    outf.write(_mail)
                    outf.close
            if _t % 1000 == 0:
                print(f"Processed {_t} mails from total {len(notesunids)}")

    print("DONE\n")

if __name__=='__main__':
   main()