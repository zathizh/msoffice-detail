#! /usr/bin/python
import docx, sys, os

if len(sys.argv) !=  2:
        print("[*] usage : " + sys.argv[0] + " <filename>")
        exit()

_ARGV1 = sys.argv[1]

if os.path.isfile(_ARGV1):
    _FILENAME = _ARGV1
    try:
        document = docx.Document(docx = _FILENAME)
        core_properties = document.core_properties
        print("[+]", end=" ")
        print(_FILENAME, end="\t: ")
        print(core_properties.author, end=": ")
#        print(time.ctime(os.path.getmtime(_FILENAME))
#        print(core_properties.created)
#        print(core_properties.last_modified_by)
#        print(core_properties.last_printed)
        print(core_properties.modified)
#        print(core_properties.revision)
#        print(core_properties.title)
#        print(core_properties.category)
#        print(core_properties.comments)
#        print(core_properties.identifier)
#        print(core_properties.keywords)
#        print(core_properties.language)
#        print(core_properties.subject)
#        print(core_properties.version)
#        print(core_properties.keywords)
#        print(core_properties.content_status)

    except Exception as ex:
        pass

else:
    if os.path.isdir(_ARGV1):
        _DIRNAME = _ARGV1
        os.chdir(_DIRNAME)
        for _FILE in os.listdir():
            try:
                document = docx.Document(docx = _FILE)
                core_properties = document.core_properties
                print("[+]", end=" ")
                print(_FILE, end="\t: ")
                print(core_properties.author, end="\t :")
#                print(core_properties.created)
#                print(core_properties.last_modified_by)
#                print(core_properties.last_printed)
                print(core_properties.modified)
#                print(core_properties.revision)
#                print(core_properties.title)
#                print(core_properties.category)
#                print(core_properties.comments)
#                print(core_properties.identifier)
#                print(core_properties.keywords)
#                print(core_properties.language)
#                print(core_properties.subject)
#                print(core_properties.version)
#                print(core_properties.keywords)
#                print(core_properties.content_status)

            except Exception as ex:
                pass
