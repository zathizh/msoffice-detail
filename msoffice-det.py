#! /usr/bin/python3
import sys, os, zipfile, lxml.etree

if len(sys.argv) !=  2:
        print("[*] usage : " + sys.argv[0] + " <filename>")
        exit()

_ARGV1 = sys.argv[1]

if os.path.isfile(_ARGV1):
    _FILENAME = _ARGV1
    try:
        zf = zipfile.ZipFile(_FILENAME)
        doc = lxml.etree.fromstring(zf.read('docProps/core.xml'))
        dc={'dc': 'http://purl.org/dc/elements/1.1/'}
        dcterms = {'dcterms': 'http://purl.org/dc/terms/'}

        creator = doc.xpath('//dc:creator', namespaces=dc)[0].text
        modified = doc.xpath('//dcterms:modified', namespaces=dcterms)[0].text[:-1].replace("T", " ")
        print(_FILENAME, end="\t: ")
        print(creator, end=" : ")
        print(modified)

    except Exception as ex:
        pass

else:
    if os.path.isdir(_ARGV1):
        _DIRNAME = _ARGV1
        os.chdir(_DIRNAME)
        for _FILENAME in os.listdir():
            try:
                zf = zipfile.ZipFile(_FILENAME)
                doc = lxml.etree.fromstring(zf.read('docProps/core.xml'))
                dc = {'dc': 'http://purl.org/dc/elements/1.1/'}
                dcterms = {'dcterms': 'http://purl.org/dc/terms/'}
                creator = doc.xpath('//dc:creator', namespaces=dc)[0].text
                modified = doc.xpath('//dcterms:modified', namespaces=dcterms)[0].text[:-1].replace("T", " ")
                print(_FILENAME, end="\t: ")
                print(creator, end=" : ")
                print(modified)

            except Exception as ex:
                pass
