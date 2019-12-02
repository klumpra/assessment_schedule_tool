'''
This program produces a page that shows "Program Assessment Progress".
The page lists Program, Contact Name, Contact Email, Outcome, and the
result of a completed assessment or schedule of upcoming ones for the
years 2019 through 2024. 

The inputs to this are 
(1) slos.xlsx, which is the export of Report #003 from TK20. 
(2) contacts.xlsx, which is the xls version of the google spreadsheet chairs
filled out to name contacts from the various programs in coast. This had been
hosted on Google Drive as sheet assessment_contacts. It has fields
Program     Assessment Contact    Assessment Contact Email
and simply lists the contact associated with each program.
(3) schedule.xls, which is actually initially produced by this program when 
INITIAL_PRODUCTION is set to True. When INITIAL_PRODUCTION is True, This 
program produce a file called assessment_schedule.xls by combining data
from contacts.xlsx and slos.xlsx and adding columns for years 2019-2024. The
resulting file looks like this:
    Contact Name    Contact Email    Student Learning Outcome    2019-2020   
    2020-2021   2021-2022    2022-2023    2023-2024
You then copy this assessment_schedule.xls to schedule.xls so that you can
format it nicely and share it with CoAST Chairs to put X's in the years when
particular outcomes will be measured. Perhaps post schedule.xls to Google 
Drive so that chairs can edit it. Or, request the chairs to specify the
schedule and update schedule.xls for them. 

This program is able to produce a web page that shows the results and upcoming
schedule of assessments for each SLO. It is informed first by the data in
TK20, as slos.xlsx (which, again, was exported from TK20) is the first source
of information for the years columns, and second by the schedule.xlsx file,
which the CoAST chairs will edit to mark when various assessments will be
done.

In this first year of the program, 2019-2020, only the 2019-2020 column
will be filled in with data read from TK20, as no future years of data are
available yet. In subsequent years, this program will have to input multiple 
slos.xlsx, each one focusing on a particular year. The "finding" variable  
represents what was found in this current year. In 2020, finding will have
to be used twice, once for 2019-2020 and again for 2020-2021, as the slos.xls
will be imported for both years. In years after that, additional slos.xlsx will 
have to be imported, and again the finding variable will be used to fill those
columns.

So, basically, the schedule is filled in first by data from TK20, which 
identifies all assessments that have been done or are in the process of being
done. Then, the input from the Chairs is read in from schedule.xlsx, which
charts the schedule forward for assessment of each student learning outcome.
'''

import xlrd
import xlwt
import datetime

'''
this function is used to turn slos into valid html anchors by removing 
punctuation and replacing spaces with hyphens. This is identical to
link_001_report.py's replace_punctuation function, except this one also 
replaces spaces with hyphens.
'''
def replace_punctuation_and_hyphenate(text):
    text = text.replace(" ","-")
    result = ""
    for ch in text:
        code = ord(ch)
        if (code >=48 and code<58) or (code >= 65 and code <=90) or (code >=97 and code <= 122):
            result = result + ch
        elif ch == "-":
            result = result + ch
    return result

INITIAL_PRODUCTION = False  
''' set this to True to create assessment_schedule.xls
                                using this program. Really, once this is done,
                                it probably never needs to be done again. So,
                                you can pretty much leave this as False from
                                here on out. But, I set it to True initially
                                so that I could produce the basis of schedule.xlsx,
                                which the chairs will fill out with the 
                                assessment schedule '''

'''
This function takes in the already open file and a list of fully formatted
tr's. Each tr contains td's that specify program contact name, contact email,
outcome, and entries for years 2019-2020 through 2023-2024.

20191130: added no_slo_programs list parameter to identify programs that
are missing SLOs in TK20.
'''
def write_output_html(fout,table_content,no_slo_programs):
    fout.write("<html>")
    fout.write("<head>")
    fout.write("<title>2019-2020 Assessment Progress</title>")
    fout.write("<style>.centered {text-align:center;} .title {font-size:36px;} .subtitle {font-size:18px;} th {background-color:DDDDDD; border: solid 1px; padding: 3px 3px 3px 3px;}")
    fout.write("td {vertical-align:top; border: solid 1px; padding-left: 3px;}")
    fout.write("a:link, a:visited {color: inherit; text-decoration:none;}")
    fout.write("</style>")
    fout.write("</head>")
    fout.write('<body style="font-family:Arial;">')
    fout.write('<center><span class="title">Program Assessment Progress</span></center><br/>')
    the_date = datetime.datetime.now()
    fout.write('<center><span class="subtitle">Report run on %s at %s</span></center><br/><br/>' % (the_date.strftime("%x"),the_date.strftime("%X")))
    fout.write("<table>")
    fout.write("<tr><th>Program</th><th>Contact Name</th><th>Contact Email</th><th>Outcome</th><th>2019-2020</th><th>2020-2021</th><th>2021-2022</th><th>2022-2023</th><th>2023-2024</th</tr>")
    for row in table_content:
        fout.write("%s\n" % row)
    fout.write("</table>")
    fout.write("<br/><br/>")
    fout.write('<p class="subtitle"><b>The following programs have no SLOs defined:</b></p>')
    fout.write("<ul>")
    for nsp in no_slo_programs:
        fout.write("<li>%s</li>" % nsp)
    fout.write("</ul>")
    fout.write("<br/><br/>")
    fout.write("</body>")
    fout.write("<html>")
    fout.close()

# The folder in which all the input and output files are located.
FOLDER = "C:\\Users\\klumpra\\Dropbox\\coast\\ray_stuff\\assessment\\"

# The output xlsx from TK20's Report #003
program_fname =  FOLDER + "slos.xlsx"

# The Excel download of the assessment_contacts google sheet the chairs did
contact_fname = FOLDER + "contacts.xlsx"

# The formatted version of assessment_schedule.xls that chairs or assessment
# contacts fill out with X's to indicate in which years various outcomes will
# be assessed. It is the more nicely formatted version of assessment_schedule.xls,
# which this program produces when INITIAL_PRODUCTION is True
schedule_fname = FOLDER + "schedule.xls"

# The excel file produced by this program when INITIAL_PRODUCTION is True.
# It is the starting point for the more nicely formatted schedule.xlsx that
# the chairs fill out. It shows the schedule for when each outcome will be
# assessed.
excel_out_name = FOLDER + "assessment_schedule.xls"

# The output file - an html file that shows the results and upcoming 
# schedule of each student learning outcome.
output_fname = FOLDER + "2019_assessment.html"

# open the output html file that will show the results and upcoming schedule
# for each student learning outcome.
fout = open(output_fname,"w")

# prog_wbk is the slos reported from TK20. We might have to have additional
# versions of this in future years when there are additional years of TK20
# reports.
prog_wbk = xlrd.open_workbook(program_fname, on_demand = True)

# cont_wbk is the data provided by chairs on who is responsibil for each
# program's assessment. They completed this activity online, and I downloaded
# it as an Excel spreadsheet.
cont_wbk = xlrd.open_workbook(contact_fname, on_demand = True)

# schd_wbk is the result of formatting assessment_schedule.xls, which
# was first produced by this program when INITIAL_PRODUCTION was set to True.
schd_wbk = xlrd.open_workbook(schedule_fname, on_demand = True)

''' If INITIAL_PRODUCTION is True, which might not every happen again,
    then go ahead and start writing the spreadsheet that specifies the 
    upcoming schedule of assessments. This will yield assessment_schedule.xls,
    which will eventually be turned into the more nicely formatted 
    schedule.xlsx that chairs and assessment contacts will edit. '''
if INITIAL_PRODUCTION:
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("schedule")
    sheet.write(0,0,"Program")
    sheet.write(0,1,"Contact Name")
    sheet.write(0,2,"Contact Email")
    sheet.write(0,3,"Student Learning Outcome")
    sheet.write(0,4,"2019-2020")
    sheet.write(0,5,"2020-2021")
    sheet.write(0,6,"2021-2022")
    sheet.write(0,7,"2022-2023")
    sheet.write(0,8,"2023-2024")
    
''' contacts a dictionary whose key is the program name and whose value
is an array. Position [0] is the name of the contact, and position [1] is
the email address  '''
contacts = {} 
cont_wks = cont_wbk.sheet_by_index(0)   # the input spreadsheet for contacts
''' for each program, read the name and email address of the contact. The key
is the program name, and the value is an array with two slots, one for contact
name and the next for contact email. '''
for row in range(2,cont_wks.nrows):  
    contacts[cont_wks.cell_value(row,0)] = [cont_wks.cell_value(row,1), cont_wks.cell_value(row,2)]

''' a dictionary whose key is the program name_SLO and whose value
is an array. [0] is 2019-2020, [1] is 2020-2021, through [4] is 2023-2024. Each
of those year slots will be either blank, have an outcome ("Met" or "Not Met") 
or have an "X" indicating that it is scheduled. '''
schedule = {}
schd_wks = schd_wbk.sheet_by_index(0)   # the input spreadsheet for schedule
for row in range(2,schd_wks.nrows):
    if schd_wks.cell_value(row,0).strip() != "":
        key = "%s_%s" % (schd_wks.cell_value(row,0).strip(),schd_wks.cell_value(row,3).strip())
        schedule[key] = [schd_wks.cell_value(row,4).strip(),schd_wks.cell_value(row,5).strip(),schd_wks.cell_value(row,6).strip(),schd_wks.cell_value(row,7).strip(),schd_wks.cell_value(row,8).strip()]

prog_wks = prog_wbk.sheet_by_index(0)   
''' the input TK20 report showing SLO and outcome
The format of this excel sheet is a bit irregular and requires moving forward
through the file to skip over blank space where a figure will appear. '''
slos = {}
i = 10
results = []
prog_num = 0
row_count = 0
'''tk20_list_of_programs hold the list of programs that were reported in the
tk20 report #003. I am keeping a list of them so that I can compare with the
keys in the contacts list to see if there are any programs that don't have
SLOs defined'''
tk20_list_of_programs = []
while i < prog_wks.nrows:
    row_count = row_count + 1
    prog_name = prog_wks.cell_value(i,0).strip().replace("  "," ") 
    ''' helps us determine the
                        corresponding contact from the contacts dictionary,
                        as the key for contacts is the program name '''
    tk20_list_of_programs.append(prog_name)
    prog_num = prog_num+1   
    ''' prog_num is used to keep track of whether to 
                        gray-scale a sequence of rows or not. This helps set 
                        aside the student learning outcomes of one program 
                        from the next in the report '''
    if prog_num % 2 == 0:   # gray-scale the even-numbered programs
        td = '<td style="background-color:#DDDDDD;">'
        tdc = '<td style="text-align:center; background-color:#DDDDDD;">'
    else:                  # don't gray-scale the odd-numbered programs
        td = "<td>"
        tdc = '<td style="text-align:center;">'
    contact_name = contacts[prog_name][0]
    contact_email = contacts[prog_name][1]
    i = i + 2  # skip over the program header to where the first SLO is
    outcome = prog_wks.cell_value(i,1).strip().replace("  "," ")
    while (outcome != ""):   
        ''' while the row contains an SLO and isn't the
                       the beginning of blank space that separates programs'''
        finding = prog_wks.cell_value(i,2).strip() 
        ''' finding fills the 2019-2020
        column. But, in future years, when data for 2020-2021 is available in TK20
        too, we'll have to do something identical for that year. finding's value
        takes precedence over what is in schedule.xls, since schedule.xls is
        supposed to just show the schedule of future assessment efforts (it uses
        X's to mark years in which each outcome will be measured). So, if the
        year is current or past, we can fill that field in with the outcome
        read from TK20. If there is no outcome from TK20, then we will grab
        whatever the value from schedule is, which will be either an X or a
        blank. An X will indicate that it's scheduled for that year.'''
        if finding == "NA":
            finding = ""
        if finding == "":
            finding = schedule["%s_%s" % (prog_name,outcome)][0]
        ''' future years aren't represented in the TK20 sheet, so we just
        grab it from the schedule sheet. In 2019-2020, there is just one 
        TK20 sheet. But, when we run this in 2020-2021, there will be a second
        TK20 sheet, and thus a need to treat f2020 the same way we treat
        finding above - by looking it up in prog_wks first (to grab the
        TK20 value) and then looking to schedule (then in position [1] instead
        of position [0]) if the finding was found to be NA)'''
        f2020 = schedule["%s_%s" % (prog_name,outcome)][1]
        f2021 = schedule["%s_%s" % (prog_name,outcome)][2]
        f2022 = schedule["%s_%s" % (prog_name,outcome)][3]
        f2023 = schedule["%s_%s" % (prog_name,outcome)][4]
        ''' html_link is the prog_name-outcome with punctuation replaced in 
        outcome and spaces replaced by hyphens. These locate places in 
        report_with_links.html, which is the exported TK20 report #001 with
        links added for programs and slo's. Those links were added by the
        program link_001_report.py. The name of the report_with_links.html
        file may have to be distinguished by year in the future.'''
        program_html_link = ("./report_with_links.html#%s" % prog_name.replace(" ","-")).strip()
        slo_html_link = ("./report_with_links.html#%s-%s" % (prog_name.replace(" ","-"),replace_punctuation_and_hyphenate(outcome))).strip()
        ''' construct a full table row with all the necessary columns, 
        including program name, contact name, contact email, SLO, and the
        year-by-year results or X's that indicate that the SLO is scheduled
        for that year. '''
        results.append('<tr>%s<a href="%s" target="_blank">%s</a></td>%s%s</td>%s%s</td>%s<a href="%s">%s</a></td>%s%s</td>%s%s</td>%s%s</td>%s%s</td>%s%s</tr>' % (td,program_html_link,prog_name,td,contact_name,td,contact_email,td,slo_html_link,outcome,tdc,finding,tdc,f2020,tdc,f2021,tdc,f2022,tdc,f2023))
        ''' this code was the one-time-only step needed to produce
        assessment_schedule.xls. From this, we'll eventually produce
        schedule.xlsx, which is the same data but more nicely formatted '''
        if INITIAL_PRODUCTION:
            sheet.write(row_count,0,prog_name)
            sheet.write(row_count,1,contact_name)
            sheet.write(row_count,2,contact_email)
            sheet.write(row_count,3,outcome)
            sheet.write(row_count,4,finding)
        row_count = row_count+1
        i=i+1
        ''' read the next SLO if there is still room in the spreadsheet '''
        if i >= prog_wks.nrows:
            outcome = ""
        else:
            outcome = prog_wks.cell_value(i,1).strip()
    ''' skip to the next program '''
    while (i < prog_wks.nrows and prog_wks.cell_value(i,0).strip() == ""):
        i = i + 1

'''determine those programs that don't have SLOs listed yet. These are ones
for which a contact has been defined but for which no SLOs appear in TK20'''        
no_slo_programs = []
for key in contacts:
    if key not in tk20_list_of_programs:
        no_slo_programs.append(key)

''' write the results as an html file. fout corresponds to output_fname '''
write_output_html(fout,results,no_slo_programs)


''' this code was used initially to produce assessment_schedule.xls. It probably
never needs to be run again.'''
if INITIAL_PRODUCTION:
    workbook.save(excel_out_name)
    workbook.release_resources()
    del workbook

prog_wbk.release_resources()
schd_wbk.release_resources()
cont_wbk.release_resources()
del prog_wbk
del schd_wbk
del cont_wbk

