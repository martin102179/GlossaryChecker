__author__ = 'machung'
# -*- coding: utf-8 -*-
#V1.2

import codecs
import csv
from operator import itemgetter
import copy
import re
import sys
import cmd
import time
import glob
from chardet.universaldetector import UniversalDetector
import pydocx
from html.parser import HTMLParser
import os


class MyParser(HTMLParser):
    def __init__(self):
        HTMLParser.__init__(self)
        self.reset()
        self.HTMLDATA = []

    def handle_data(self, data):
        self.HTMLDATA.append(data)

    def clean(self):
        self.HTMLDATA = []

alphabet = set(['a','b','c','d','e','f','g','h','i','j','k''l','m','n','o','p','q','r','s','t','u','v','w','x','y','z','A','B',
                     'C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'])
back_alphabet = set(['a','b','c','d','e','f,''g','h','i','j','k''l','m','n','o','p','q','r','t','u','v','w','x','y','z','A','B',
                    'C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','T','U','V','W','X','Y','Z'])

htmlhead_one = '''<style>
table, th, td {
border: 1px solid black;
border-collapse: collapse;
padding: 5px;
text-align: left;
}

a.correct {
background-color: lightgreen;
}

span.incorrect {
background-color: ff5c33;
}

body{
margin: 0;
padding: 0;
border: 0;
overflow: hidden;
height: 100%;
max-height: 100%;
}

#framecontentleft{
position: absolute;
top: 0;
left: 0;
width: 100%;
height: 130px; /*Height of frame div*/
overflow: hidden; /*Disable scrollbars. Set to "scroll" to enable*/
background-color: navy;
color: white;
}

#framecontentright{
position: absolute;
top: 0;
left: 600;
width: 50%;
height: 130px; /*Height of frame div*/
overflow: hidden; /*Disable scrollbars. Set to "scroll" to enable*/
background-color: navy;
color: white;
}

#maincontent{
position: fixed;
top: 130px; /*Set top value to HeightOfFrameDiv*/
left: 0;
right: 0;
bottom: 0;
overflow: auto;
background: #fff;
}

.innertube{
margin: 15px; /*Margins for inner DIV inside each DIV (to provide padding)*/
}
</style>
<script>
function rearrange(){
    if (HideTag.checked==1){
        HideTagFunction();
    }else{
        ShowTagFunction();
    }
    if (MinorHide.checked==1){
        MinorHideFunction();
    }else{
        MinorShowFunction();
    }
    if (NoErrorhide.checked==1){
        NoErrorhideFunction();
    }else{
        NoErrorshowFunction();
    }
}

function NoErrorhideFunction() {
    $(".NoError").hide();
}

function NoErrorshowFunction() {
    $(".NoError").show();
}

function HideTagFunction() {
    $("a").removeClass("correct");
}

function ShowTagFunction() {
    $("a").addClass("correct");
}

function MinorHideFunction() {
    $(".minor").hide();
}

function MinorShowFunction() {
    $(".minor").show();
}

window.onload = NoErrorhideFunction</script>

<body>
<div id="framecontentleft">
<div class="innertube">
<script src="jquery-3.1.0.js"></script>
<input type="checkbox" id="HideTag" name="HideTag">Unmark correct glossaries<br>
<input type="checkbox" id="MinorHide" name="MinorHide">Hide lines that only has minor mistake ( {} tag in glossary comment)<br>
<input type="checkbox" id="NoErrorhide" name="NoErrorhide" checked>Hide lines with no glossary error<br>
<button onclick="rearrange()">Reload</button>
<div id="framecontentright">
<div class="innertube">'''

htmlhead_two='Segments in source file: %d<br>'
htmlhead_three='Segments in localized file: %d<br>'
htmlhead_four='''</div>
</div>
</div>
</div>

<div id="maincontent">
<div class="innertube">
<table style="width:100%;table-layout:fixed">'''



def Read_in_Glossary(Path):
    try:
        ProcessedList=[]
        fin = codecs.open('%s' % Path, 'r', encoding='utf-8', errors='ignore')
        while True:
            line =fin.readline()
            if not line:
                break
            temp =line.split('\t')
            if len(temp) == 4:
                temp.pop(0)
                ProcessedList.append(temp)
            else:
                ProcessedList[len(ProcessedList)-1][2] += temp[0]
        fin.close()
        ProcessedList.pop(0)
        ProcessedList.append(['error-empty', 'The enUS version has less segment than localized file.', ''])
        #Generate_Log('Glossary.txt is loaded successfully.\n')
    except:
        print('Glossary.txt is not found.')
        return []
        #Generate_Log('Glossary.txt is not found.\n')
    return ProcessedList


def Parse_Document(html):

    #*******************************************************************************************************************
    # The following .replace function inserts a special character before each tag that is likely to represent a text
    # segment. Then use the .split function to split these segments into a list. Source file and target file should have
    # the same number of tags. So the list index should be equal between the source and the target list.
    #*******************************************************************************************************************
    html = html.replace('<p', '※<p' )
    html = html.replace('<h', '※<h')
    #html = html.replace('<span style="font', '※<span style="font')
    html = html.replace('<block', '※<block')
    html = html.replace('<table', '※<table')
    html = html.replace('<li', '※<li')
    # replace \ufffd replacement character with single quote.
    html = html.replace(u'\ufffd', '\'')
    html = html.replace('’', '\'')
    #html = html.replace(' \' ', '－')
    html = html.replace('&mdash;' , '')
    html = html.replace('&nbsp;' , '')
    html = html.replace('&#160;' , '')

    #Split HTML at where '※' is inserted
    TempList = html.split('※')
    ProcessedList=[]
    for i in range(len(TempList)):

        Parser = MyParser()
        Parser.feed(TempList[i])
        textcontent = ''.join(Parser.HTMLDATA)
        Parser.clean()
        # removing lines that has 'blockquote'  and  'pydocx-caps'  which usually should not be localized.
        if 'blockquote' not in textcontent and 'pydocx-caps' not in textcontent:
            #lines should not start with \r and should not be empty
            if (len(textcontent) > 0) and (textcontent[0:1] != '\r') and (textcontent[0:2] != ' \r'):
                #lines should not start as single quote '
                if (textcontent[0:1] != '\''):
                    ProcessedList.append(textcontent)

    return ProcessedList


def Check_HTML_Codepage(Path):
    #HTML file encoding detector
    try:
        detector = UniversalDetector()
        detector.reset()
        textdetect = open(Path, 'rb')
        detector.feed(textdetect.read())
        detector.close()
        print(detector.result)
    except:
        return None
    return detector.result['encoding']



def Read_in_File(Path, mode):
    # Path provide file path
    # mode decides if this is a HTML or a Docx
    # if_this_is_html_source_file needs to be assigne True when reading the source HTML file. Otherwise it can be
    # ignored.

    docxreader = pydocx.PyDocX

    try:
        if mode == 'html':
            #if 'source' in Path:
                # if this is HTML source file, open it in cp1252
                #fin = codecs.open('%s' % Path, 'r', encoding='cp1252', errors='replace')
            #else:
            fin = codecs.open('%s' % Path, 'r', encoding='utf-8', errors='replace')
            temp = fin.read()
            fin.close()

        elif mode == 'docx':
            temp = docxreader.to_html(Path)
        print('%s found' % Path)
        # read file success, return a parsed document (it's a list)
        return Parse_Document(temp)

    except IOError:
        print('%s not found' % Path)
        # read file failed. return an empty list
        return []



def Glossary_Check(SrcTable, TargetTable, Glossary):
    def Change_Column_To_Text(matrix, i):
        # Helper function to join text from a list column
        temp = [row[i] for row in matrix]
        return ','.join(temp)


    # ScrTable VS TargetTable should have the same number of lines. If not, '*****error-empty*****' is appended to which ever has less number of lines.
    # until the lines are equal.
    if len(SrcTable) > len(TargetTable):
        number = len(SrcTable)
        while len(SrcTable) != len(TargetTable):
            TargetTable.append('*****error-empty*****')
    elif len(SrcTable) < len(TargetTable):
        number = len(TargetTable)
        while len(SrcTable) != len(TargetTable):
            SrcTable.append('*****error-empty*****')

    if len(SrcTable) == len(TargetTable):
        number = len(TargetTable)

        OutputTable =[]
        OutputTable.append(['Source', 'Target', 'Glossary and Comments', ''])

        for g in range(number):
        # for each glossary found in the Source text, there must be a matching localized translation in the target text.
        # Use NotMatchedGlossary to store these glossaries we found. Append glossary length column since we want to sort the findings according to length
        # Use MatchedGlossary to store the matched glossaries we found. Also, Append glossary length column.
            MatchedGlossary = []
            NotMatchedGlossary = []
            for y in range(len(Glossary)):
                CopyList =[]
                if str(Glossary[y][0]).lower() in str(SrcTable[g]).lower():
                    if Glossary[y][1] in TargetTable[g]:
                        CopyList = copy.deepcopy(Glossary[y])
                        CopyList.append(len(Glossary[y][0]))
                        MatchedGlossary.append(CopyList)
                    if Glossary[y][1] not in TargetTable[g]:
                        CopyList = copy.deepcopy(Glossary[y])
                        CopyList.append(len(Glossary[y][0]))
                        NotMatchedGlossary.append(CopyList)

            #**********************************************************************************************************
            # After all the matched Glossary or not matched Glossary is found. We will sort them by the length of the
            # Glossary. We want to deal with the longest glossaries first than the shorter ones.
            #*********************************************************************************************************
            NotMatchedGlossary = sorted(NotMatchedGlossary, key=itemgetter(3), reverse=True)
            MatchedGlossary = sorted(MatchedGlossary, key=itemgetter(3), reverse=True)



            #call Markup_Up_Strings function to mark up glossary in the source and target text.
            matched_glossary_set = Change_Column_To_Text(MatchedGlossary, 0)
            notmatched_glossary_set = Change_Column_To_Text(NotMatchedGlossary, 0)


            SourceTextList = [SrcTable[g]]
            TargetTextList = [TargetTable[g]]

            # Processing Matched Glossary cases first. Calling Mark_UP() to slice and mark the glossary
            for elem in range(len(MatchedGlossary)):

                # Matched Glosaary should not be in any part of the Not Matched Glossaries.
                if str(MatchedGlossary[elem][0]).upper() not in notmatched_glossary_set.upper():
                    len_before_markup = len(SourceTextList)

                    SourceTextList = Mark_Up(str(MatchedGlossary[elem][0]), SourceTextList, 'match')

                    len_after_markup = len(SourceTextList)

                    # if TextList length changed means we found an matching glossary, also mark up correct glossary in the target
                    if len_before_markup < len_after_markup:
                        TargetTextList = Mark_Up(str(MatchedGlossary[elem][1]), TargetTextList, 'match')

            comments = ''  # contains the content for glossary comment field of the report
            ifminor = ''  # flag for minor issues. Will be used as identification in Generate_Result()

            # Processing the  Not Matched Glossary cases.
            for elem in range(len(NotMatchedGlossary)):

                # Not Matched Glosaary should not be in any part of the Matched Glossaries.
                if str(NotMatchedGlossary[elem][0]).upper() not in matched_glossary_set.upper():

                    len_before_markup = len(SourceTextList)

                    SourceTextList = Mark_Up(str(NotMatchedGlossary[elem][0]), SourceTextList, 'notmatch')

                    len_after_markup = len (SourceTextList)

                    # if TextList length changed means we found an glossary error, add correct glossary with comments.
                    if len_before_markup < len_after_markup:

                        # If Glossary Comments has never been added for this line, add Glossary Comments.
                        if comments == '':
                            if len(NotMatchedGlossary[elem][2]) > 2:
                                comments = '<li>' + NotMatchedGlossary[elem][0] + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + \
                                           NotMatchedGlossary[elem][
                                               1] + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + '<i>Comment:' + \
                                           NotMatchedGlossary[elem][2] + '</i></li>'
                            else:
                                comments = '<li>' + NotMatchedGlossary[elem][0] + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + \
                                           NotMatchedGlossary[elem][
                                               1] + '</li>'

                        # If Glossary Comments has been added for this line, there are multiple glossary errors found. insert line break tags and add
                        # more comments
                        else:
                            if len(NotMatchedGlossary[elem][2]) > 2:
                                comments = comments + '<br><br><li>' + NotMatchedGlossary[elem][
                                    0] + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + \
                                           NotMatchedGlossary[elem][
                                               1] + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<i>Comment:' + \
                                           NotMatchedGlossary[elem][2] + '</i></li>'
                            else:
                                comments = comments + '<br><br><li>' + NotMatchedGlossary[elem][
                                    0] + '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;' + \
                                           NotMatchedGlossary[elem][
                                               1] + '</li>'

                        if ('{' and '}') in comments and comments.count('</li>') == comments.count('}'):
                            # if the minor error's alternative translation is in the localized html, ifminor flag is set to True. When output is generated,
                            # minor mistakes can be filtered.
                            if comments[comments.find('{') + 1:comments.find('}')] in TargetTable[g]:
                                ifminor = True
                            else:
                                ifminor = None
                        else:
                            ifminor = None

            # calling .join() method to join the content of the text hold list in to one line string.
            englishtext = str(''.join(SourceTextList))
            localizedtext = str(''.join(TargetTextList))

            if localizedtext == '*****error-empty*****':
                localizedtext = '<span class="incorrect">' + localizedtext + '</span>'
                comments = '*****Localized file has less segments than enUS version.*****'

            TempList = [englishtext,localizedtext,comments, ifminor]

            OutputTable.append(TempList)




    return OutputTable


def Mark_Up(glossary, TextList, matchflag):
    # This is to mark the correct/incorrect glossary in Scour/Target line with HTML tags
    flag = 0
    wordlength = len(glossary)

    # The while statement will go through the TextList list to check if there are glossaries inside.
    # if a glossary is found. TextList list will be sliced into 3 segments. for example, TextList = ['Cast Holylight now']
    # Holylight is recognized as glossary. The  updated TextList list will be ['Cast', '<some tag>Holylight</some tag>, 'now']
    # List elements with html tag inside are skipped to avoid marking
    while flag < len(TextList):
        string =  str(TextList[flag]).upper()

        position = string.find(glossary.upper())

        #if position < 0, the glossary is not found in the TextList list.
        if position == -1:

            flag += 1
        else:



            # To check the glossary found in the text is independant vocabulary (no alphabet appears in front of or behind it).
            # Only exception is 's'  since it may be in plural from.

            if TextList[flag][position - 1:position] not in alphabet:
                if TextList[flag][position + wordlength:position + wordlength + 1] not in back_alphabet:
                    if TextList[flag][position + wordlength:position + wordlength + 1] == ('s' or 'S') \
                            and TextList[flag][position + wordlength:position + wordlength + 2] in alphabet:

                        flag += 1



                    elif matchflag == 'match':
                        # Mark up matched glossary by inserting <a class="correct"></a> tag if the list element has not been marked up yet.
                        if '</a>' not in TextList[flag]:

                            #copy the content from the position where glossary is found up to the end of the element.
                            temp = TextList[flag][position:]

                            #slice the current element up from the begining up to the the position where the glossary is found
                            TextList[flag] = TextList[flag][:position]
                            if TextList[flag] == '':
                                TextList[flag] = '<a class="correct">' + temp[:wordlength] + '</a>'
                                TextList.insert(flag + 1, temp[wordlength:])
                            else:
                                # insert a list element that contains our glossary plus the <a class...  tags
                                TextList.insert(flag + 1, '<a class="correct">' + temp[:wordlength] + '</a>')

                                # insert a list element  that
                                TextList.insert(flag + 2, temp[wordlength:])

                        else:
                            flag += 1

                    else:
                        # Mark up unmatched glossary with <span class="incorrect"></span> tags if it has not been marked up yet.

                        if ('</a>' and '</span>') not in TextList[flag]:
                            temp = TextList[flag][position:]
                            TextList[flag] = TextList[flag][:position]
                            if TextList[flag] =='':
                                TextList[flag] = '<span class="incorrect">' + temp[:wordlength] + '</span>'
                                TextList.insert(flag + 1, temp[wordlength:])
                            else:
                                TextList.insert(flag + 1, '<span class="incorrect">' +
                                                    temp[:wordlength] + '</span>')
                                TextList.insert(flag + 2, temp[wordlength:])
                        else:
                            flag += 1

                else:
                    flag += 1

            else:
                flag += 1

    return TextList


def Generate_Result(OutputTable, mode, enlength, targetlength, targetencoding=None):
    # Function to be results into a HTML table. targetencoding variable must be provided if the document is HTML.
    # targetencoding is not necessary parameter when processing other file format.

    try:
        os.mkdir("Output")
    except:
        pass
    try:
        if mode == 'html':
            fout = open(r"Output\HTML_check_Result.html", "w",encoding='utf-8')
        if mode == 'docx':
            fout = open(r"Output\DOCX_check_Result.html", "w", encoding='utf-8')

        fout.write(u'\ufeff') #writing UTF8 BOM to ensure the output is in correct encoding

        #writing html file header with CSS/Javascript
        fout.write(htmlhead_one)

        #insert  addition info into
        fout.write('Segments in source file:    %d' % (enlength) + '<br>')
        if mode == 'html':
            fout.write(
                'Segments in localized file: %d' % (targetlength) + '<br>' + 'Localized HTML Codepage: %s<br>' % (
                targetencoding))

        elif mode =='docx':
            fout.write(
                'Segments in localized file: %d' % (targetlength) + '<br><br>' )

        if targetencoding != None and ('UTF-8' not in targetencoding.upper()):
            fout.write('<b><font color ="red">Warning: Glossary Checker '
                       'found that localized version is not encoded in UTF-8. The check result will be inaccurate. '
                       'Please re-save localized HTML in UTF-8 encoding.</font><b>')

        elif enlength > targetlength:
            fout.write('<b><font color ="red">Warning: '
                'Glossary Checker found that enUS version has more segments than localized'
                ' version. Please double check localized file to make sure that you did not delete any '
                       'HTML tags or Word paragraphs.</font><b>')

        elif targetlength > enlength:
            fout.write('<b><font color ="red">Warning: '
                'Glossary Checker found that localized version has more segments than enUS'
                ' version. Please double check localized file to make sure that you did not insert extra '
                       'HTML tags or Word paragraphs.</font><b>')

        #writing the rest of the html header
        fout.write(htmlhead_four)

        # build tables for Glossary Check result
        for z in range(len(OutputTable)):
            
            #mark lines only has minor issue. insert <tr class="minor>
            if OutputTable[z][3] == True:
                fout.write('<tr class="minor"><td style="width:30px">' +str(z) + '</td><td>' + OutputTable[z][0] + '</td>' + '<td>' + OutputTable[z][1] +
                    '</td>' + '<td>' + OutputTable[z][2] + '</td></tr>\n')

            # mark line that are all correct. insert <tr class="NoError">
            elif OutputTable[z][2] == '':
                fout.write('<tr class="NoError"><td style="width:30px">' +str(z) + '</td><td>' + OutputTable[z][0] + '</td>' + '<td>' + OutputTable[z][
                    1] + '</td>' + '<td>' + OutputTable[z][2] + '</td></tr>\n')
            
            #below deals with line that has error.
            else:
                fout.write('<tr><td style="width:30px">' +str(z) + '</td><td>' + OutputTable[z][0] + '</td>' + '<td>' + OutputTable[z][1] +
                    '</td>' + '<td>' + OutputTable[z][2] + '</td></tr>\n')

        fout.write('</table>\n')
        fout.write('</body>')
        fout.close()
        print('%s Glossary Check completed\n%s_Check_Result.html is saved successfully\n' % (mode.upper(), mode))
        return 1

    except IOError:
        print ('Error：Unable to write %s_Check_Result.html\n' % mode)
        return 0






def main():
    #Generate_Log('new log') # empty the log file
    GlossaryList = Read_in_Glossary(r'Glossary.txt')
    if GlossaryList != []:
        SrcList = []
        TargetList = []
        OutPutList = []
        html_source_path = r'source.html'
        html_target_path = r'target.html'
        DOCXsource_path = r'source.docx'
        DOCXtarget_path = r'target.docx'

        # Read in Html files

        SrcList = Read_in_File(html_source_path, 'html')
        TargetList = Read_in_File(html_target_path, 'html')
        targetencoding = Check_HTML_Codepage(html_target_path)

        # Processing Html files

        enlength = len(SrcList)
        targetlength = len(TargetList)
        OutPutList = Glossary_Check(SrcList, TargetList, GlossaryList, )
        result = Generate_Result(OutPutList, 'html', enlength, targetlength, targetencoding)

        # Read in Docx files
        SrcList = Read_in_File(DOCXsource_path, 'docx')
        TargetList = Read_in_File(DOCXtarget_path, 'docx')

        # Processing Docx files

        enlength = len(SrcList)
        targetlength = len(TargetList)
        OutPutList = Glossary_Check(SrcList, TargetList, GlossaryList)
        result = Generate_Result(OutPutList, 'docx', enlength, targetlength)
        time.sleep(2)
    else:
        print('Glossary.txt required for glossary check. Exiting Glossary Checker.')

if __name__ == '__main__':
    main()
    time.sleep(5)





