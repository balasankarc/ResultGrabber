import json
import operator
import collections
mapping = [
    'Principles of Management',
    'Engineering Mathematics IV',
    'Database Management Systems',
    'Digital Signal Processing',
    'Operating Systems',
    'Advanced Microprocessors and Peripherals',
    'Database Lab',
    'Hardware and Microprocessors Lab'
]
f = open('output.json')
content = f.read()
result1 = json.loads(content)["Adi Shankara Institute Of Engineering And Technology"][
    "Computer Science And Engineering"]
result = collections.OrderedDict(result1)
outputstring = ""
for register, details in result.items():
    outputstring += register
    name = details["name"]
    details.pop("name")
    outputstring += "," + name
    for subject in mapping:
        for mark in details[subject]:
            outputstring += "," + str(mark)
    outputstring += "\n"
print outputstring
