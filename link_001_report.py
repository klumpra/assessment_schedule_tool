'''
This program takes the TK20 001 Report exported as html and inserts anchors in it
so that the assessment report produced by assessment_schedule.py can link to
parts of the report. Each program is anchored, and each SLO within each 
program is anchored. This allows the assessment schedule report to link to each
program and to each outcome listed in the 001 report.

Note: do not use main.css. Just use the bootstrap css. Otherwise, the anchors
aren't handled properly - the page doesn't scroll to the anchor because the
main.css has too many floating elements that end up screwing up the scrolling.
'''

'''
This function removes punctuation from the SLOs. It will preserve hyphens,
since hyphens will have replaced spaces by the time this function is called.
'''
def remove_punctuation(text):
    result = ""
    for ch in text:
        code = ord(ch)
        if (code >=48 and code<58) or (code >= 65 and code <=90) or (code >=97 and code <= 122):
            result = result + ch
        elif ch == "-":
            result = result + ch
    return result
            
'''
The TK20 001 report was downloaded as report.html. From this, we create
report_with_links.html.
'''

fin = open("report.html","r",encoding="utf-8")
fout = open("report_with_links.html","w",encoding="utf-8")
''' there is actually just one line of text in the file when it is exported.'''
text = fin.readlines()[0]
'''each program's SLO report begins with a div of class heading '''
parts = text.split('<div class = "heading"  >')
for i,part in enumerate(parts):
    if True:
        part = part.strip()
        if i == 0:   # the first part, before the first heading div
            fout.write(part)
        else:
            start = part.find(">")
            end = part.find("<",start)
            prog = part[start+1:end].replace(" ","-")
            part = part[:end] + "</h2>" + part[end:]
            ''' some lines have double spaces - replace with a single space '''
            part = part.replace("  "," ")
            fout.write('<div class = "heading" id = "%s" >%s' % (prog,part))
            ''' now look within each program at the individual SLOs '''
            slos = part.split('div class="max-width-text"><font size="2" style=\'text-transform: uppercase;\'>')
            slo_text = ""
            for j,slo in enumerate(slos):
                if j == 0:
                    slo_text = slo
                else:
                    end = slo.find("<")
                    slo_name = remove_punctuation(slo[:end].replace("  "," ").replace(" ","-")) 
                    slo_text = slo_text + ('div class="max-width-text" id="%s-%s"><font size="2" style=\'text-transform:uppercase;\'>' % (prog,slo_name)) + slo
    #        fout.write('<div class = "heading" >%s' % part)
    #        print('<div class = "heading" >%s' % slo_text)
            fout.write('<div class = "heading" >%s' % slo_text)
fin.close()
fout.close()
