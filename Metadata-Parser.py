# Import the os module for accessing files on the local computer.
import os
# Import the re module to permit the use of regular expressions.
import re
# Import the csv module to write to a CSV file
import csv



# Create a dictionary to provide numerical equivalents for month strings.
monthDict = {
    'Jan': '1',
    'Feb': '2',
    'Mar': '3',
    'Apr': '4',
    'May': '5',
    'Jun': '6',
    'Jul': '7',
    'Aug': '8',
    'Sep': '9',
    'Oct': '10',
    'Nov': '11',
    'Dec': '12',
}



# Create regular expressions to parse information from the file.
regExDict = {
    'Message-ID_': '^Message-ID:(.*)\n',
    'Date_From': '\nDate:((\n|.)*)((\nFrom: ))',
    'Date_To': '\nDate:((\n|.)*)((\nTo: ))',
    'Date_Subject': '\nDate:((\n|.)*)((\nSubject: ))',
    'Date_Cc': '\nDate:((\n|.)*)((\nCc: ))',
    'Date_Mime-Version': '\nDate:((\n|.)*)((\nMime-Version: ))',
    'Date_Content-Type': '\nDate:((\n|.)*)((\nContent-Type: ))',
    'Date_Content-Transfer-Encoding': '\nDate:((\n|.)*)((\nContent-Transfer-Encoding: ))',
    'Date_Bcc': '\nDate:((\n|.)*)((\nBcc: ))',
    'Date_X-From': '\nDate:((\n|.)*)((\nX-From: ))',
    'Date_X-To': '\nDate:((\n|.)*)((\nX-To: ))',
    'Date_X-cc': '\nDate:((\n|.)*)((\nX-cc: ))',
    'Date_X-bcc': '\nDate:((\n|.)*)((\nX-bcc: ))',
    'Date_X-Folder': '\nDate:((\n|.)*)((\nX-Folder: ))',
    'Date_X-Origin': '\nDate:((\n|.)*)((\nX-Origin: ))',
    'Date_X-FileName': '\nDate:((\n|.)*)((\nX-FileName: ))',
    'From_To': '\nFrom:((\n|.)*)((\nTo: ))',
    'From_Subject': '\nFrom:((\n|.)*)((\nSubject: ))',
    'From_Cc': '\nFrom:((\n|.)*)((\nCc: ))',
    'From_Mime-Version': '\nFrom:((\n|.)*)((\nMime-Version: ))',
    'From_Content-Type': '\nFrom:((\n|.)*)((\nContent-Type: ))',
    'From_Content-Transfer-Encoding': '\nFrom:((\n|.)*)((\nContent-Transfer-Encoding: ))',
    'From_Bcc': '\nFrom:((\n|.)*)((\nBcc: ))',
    'From_X-From': '\nFrom:((\n|.)*)((\nX-From: ))',
    'From_X-To': '\nFrom:((\n|.)*)((\nX-To: ))',
    'From_X-cc': '\nFrom:((\n|.)*)((\nX-cc: ))',
    'From_X-bcc': '\nFrom:((\n|.)*)((\nX-bcc: ))',
    'From_X-Folder': '\nFrom:((\n|.)*)((\nX-Folder: ))',
    'From_X-Origin': '\nFrom:((\n|.)*)((\nX-Origin: ))',
    'From_X-FileName': '\nFrom:((\n|.)*)((\nX-FileName: ))',
    'To_Subject': '\nTo:((\n|.)*)((\nSubject: ))',
    'To_Cc': '\nTo:((\n|.)*)((\nCc: ))',
    'To_Mime-Version': '\nTo:((\n|.)*)((\nMime-Version: ))',
    'To_Content-Type': '\nTo:((\n|.)*)((\nContent-Type: ))',
    'To_Content-Transfer-Encoding': '\nTo:((\n|.)*)((\nContent-Transfer-Encoding: ))',
    'To_Bcc': '\nTo:((\n|.)*)((\nBcc: ))',
    'To_X-From': '\nTo:((\n|.)*)((\nX-From: ))',
    'To_X-To': '\nTo:((\n|.)*)((\nX-To: ))',
    'To_X-cc': '\nTo:((\n|.)*)((\nX-cc: ))',
    'To_X-bcc': '\nTo:((\n|.)*)((\nX-bcc: ))',
    'To_X-Folder': '\nTo:((\n|.)*)((\nX-Folder: ))',
    'To_X-Origin': '\nTo:((\n|.)*)((\nX-Origin: ))',
    'To_X-FileName': '\nTo:((\n|.)*)((\nX-FileName: ))',
    'Subject_Cc': '\nSubject:((\n|.)*)((\nCc: ))',
    'Subject_Mime-Version': '\nSubject:((\n|.)*)((\nMime-Version: ))',
    'Subject_Content-Type': '\nSubject:((\n|.)*)((\nContent-Type: ))',
    'Subject_Content-Transfer-Encoding': '\nSubject:((\n|.)*)((\nContent-Transfer-Encoding: ))',
    'Subject_Bcc': '\nSubject:((\n|.)*)((\nBcc: ))',
    'Subject_X-From': '\nSubject:((\n|.)*)((\nX-From: ))',
    'Subject_X-To': '\nSubject:((\n|.)*)((\nX-To: ))',
    'Subject_X-cc': '\nSubject:((\n|.)*)((\nX-cc: ))',
    'Subject_X-bcc': '\nSubject:((\n|.)*)((\nX-bcc: ))',
    'Subject_X-Folder': '\nSubject:((\n|.)*)((\nX-Folder: ))',
    'Subject_X-Origin': '\nSubject:((\n|.)*)((\nX-Origin: ))',
    'Subject_X-FileName': '\nSubject:((\n|.)*)((\nX-FileName: ))',
    'Cc_Mime-Version': '\nCc:((\n|.)*)((\nMime-Version: ))',
    'Cc_Content-Type': '\nCc:((\n|.)*)((\nContent-Type: ))',
    'Cc_Content-Transfer-Encoding': '\nCc:((\n|.)*)((\nContent-Transfer-Encoding: ))',
    'Cc_Bcc': '\nCc:((\n|.)*)((\nBcc: ))',
    'Cc_X-From': '\nCc:((\n|.)*)((\nX-From: ))',
    'Cc_X-To': '\nCc:((\n|.)*)((\nX-To: ))',
    'Cc_X-cc': '\nCc:((\n|.)*)((\nX-cc: ))',
    'Cc_X-bcc': '\nCc:((\n|.)*)((\nX-bcc: ))',
    'Cc_X-Folder': '\nCc:((\n|.)*)((\nX-Folder: ))',
    'Cc_X-Origin': '\nCc:((\n|.)*)((\nX-Origin: ))',
    'Cc_X-FileName': '\nCc:((\n|.)*)((\nX-FileName: ))',
    'Bcc_X-From': '\nBcc:((\n|.)*)((\nX-From: ))',
    'Bcc_X-To': '\nBcc:((\n|.)*)((\nX-To: ))',
    'Bcc_X-cc': '\nBcc:((\n|.)*)((\nX-cc: ))',
    'Bcc_X-bcc': '\nBcc:((\n|.)*)((\nX-bcc: ))',
    'Bcc_X-Folder': '\nBcc:((\n|.)*)((\nX-Folder: ))',
    'Bcc_X-Origin': '\nBcc:((\n|.)*)((\nX-Origin: ))',
    'Bcc_X-FileName': '\nBcc:((\n|.)*)((\nX-FileName: ))',
    'X-From_X-To': '\nX-From:((\n|.)*)((\nX-To: ))',
    'X-From_X-cc': '\nX-From:((\n|.)*)((\nX-cc: ))',
    'X-From_X-bcc': '\nX-From:((\n|.)*)((\nX-bcc: ))',
    'X-From_X-Folder': '\nX-From:((\n|.)*)((\nX-Folder: ))',
    'X-From_X-Origin': '\nX-From:((\n|.)*)((\nX-Origin: ))',
    'X-From_X-FileName': '\nX-From:((\n|.)*)((\nX-FileName: ))',
    'X-To_X-cc': '\nX-To:((\n|.)*)((\nX-cc: ))',
    'X-To_X-bcc': '\nX-To:((\n|.)*)((\nX-bcc: ))',
    'X-To_X-Folder': '\nX-To:((\n|.)*)((\nX-Folder: ))',
    'X-To_X-Origin': '\nX-To:((\n|.)*)((\nX-Origin: ))',
    'X-To_X-FileName': '\nX-To:((\n|.)*)((\nX-FileName: ))',
    'X-cc_X-bcc': '\nX-cc:((\n|.)*)((\nX-bcc: ))',
    'X-cc_X-Folder': '\nX-cc:((\n|.)*)((\nX-Folder: ))',
    'X-cc_X-Origin': '\nX-cc:((\n|.)*)((\nX-Origin: ))',
    'X-cc_X-FileName': '\nX-cc:((\n|.)*)((\nX-FileName: ))',
    'X-bcc_X-Folder': '\nX-bcc:((\n|.)*)((\nX-Folder: ))',
    'X-bcc_X-Origin': '\nX-bcc:((\n|.)*)((\nX-Origin: ))',
    'X-bcc_X-FileName': '\nX-bcc:((\n|.)*)((\nX-FileName: ))',
    'X-Folder_X-Origin': '\nX-Folder:((\n|.)*)((\nX-Origin: ))',
    'X-Folder_X-FileName': '\nX-Folder:((\n|.)*)((\nX-FileName: ))',
    'X-Origin_X-FileName': 'X-Origin:((\n|.)*)((X-FileName: ))',
    'X-FileName_': 'X-FileName:(.*)\n',
    'DateParser': '(Mon|Tue|Wed|Thu|Fri|Sat|Sun), (\d.*) (Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) (\d\d\d\d) (\d\d:\d\d:\d\d)(.*) \((.*)\)'
}



# Function to return a string of the information specified by a regular expression.
def parseInfo(text, regExKey_1, regExKey_2):
    # text: The actual text (or content) to be read.
    # regExKey_1: The first half of the regex key (to call the appropriate dictionary enrty).
    # regExKey_2: The second half of the regex key.
    regExpr = regExDict[regExKey_1+'_'+regExKey_2]
    default_info = 'NA'
    info = re.findall(regExpr, text)
    if info == []:
        return default_info
    else:
        list_info = info[0]
        if(type(list_info)==type(())):
            list_info = list_info[0]
            
        list_info = list_info.replace('\n', ' ')
        list_info = list_info.replace('\t', ' ')
        list_info = list_info.strip()
        if list_info == '':
            list_info = ' '
        return list_info



count = 0

# Specify a path to the desired folder here. The folder must ONLY contain email files:
dir_path = '.../Sample Data/skilling-j/_sent_mail' # REPLACE this string with the folder path.

# method listdir() returns a list containing the names of the entries in the directory given by dir_path.
emailList = os.listdir(dir_path)

# '.DS_Store' is a file used by the Mac OS to provide information about its surrounding folder. There is no need
#    for it in this list.
if '.DS_Store' in emailList:
    emailList.remove('.DS_Store')

print('Number of Files: ', len(emailList))



with open('email_file.csv', mode='w') as email_file:
    email_writer = csv.writer(email_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    email_writer.writerow(['Message-ID', 'Date', 'Time', 'Timezone', 'From', 'To', 'Subject', 'Cc', 'Bcc', 'X-From', 'X-To', 'X-cc', 'X-bcc', 'X-Folder', 'X-Origin', 'X-FileName'])
    for specificFile in emailList:
        fileName_in = os.path.join(dir_path, specificFile)
        f = open(fileName_in, 'r')
        if f.mode == 'r':
            # Read through the file and separate the section containing the meta data.
            raw_lines = f.readlines()
            max_indx = len(raw_lines)-1
            metadata_boundary = max_indx
            for i_indx in range(len(raw_lines)):
                if 'X-FileName: ' in raw_lines[i_indx]:
                    metadata_boundary = i_indx + 1
            for a_indx in range(metadata_boundary, max_indx):
                raw_lines.pop()
            contents = ''
            for single_line in raw_lines:
                contents = contents + single_line
        
            print('-------------------------------(' + specificFile + ')-------------------------------')
            #file_info = ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA']
        
            # Here are the attributes to be imported from the text file into 'file_info':
            #   [0]: Message Identifier
            #   [1]: Date
            #   [2]: Time
            #   [3]: Timezone
            #   [4]: From (or Sender) email address
            #   [5]: To (or Receiver) email address
            #   [6]: Subject Line
            #   [7]: Cc
            #   [8]: Bcc
            #   [9]: X-From
            #   [10]: X-To
            #   [11]: X-cc
            #   [12]: X-bcc
            #   [13]: X-Folder
            #   [14]: X-Origin
            #   [15]: X-FileName
        
            # NOTE: This process begins with a list of size 14 because there are 14 fields of interest
            #       from which we are interested in parsing data from.
            field_tuple = ('Message-ID', 'Date', 'From', 'To', 'Subject', 'Cc', 'Mime-Version', 'Content-Type', 'Content-Transfer-Encoding', 'Bcc', 'X-From', 'X-To', 'X-cc', 'X-bcc', 'X-Folder', 'X-Origin', 'X-FileName')
            desired_field_tuple = ('Message-ID', 'Date', 'From', 'To', 'Subject', 'Cc', 'Bcc', 'X-From', 'X-To', 'X-cc', 'X-bcc', 'X-Folder', 'X-Origin', 'X-FileName')
            final_field_tuple = ('Message-ID', 'Date', 'Time', 'Timezone', 'From', 'To', 'Subject', 'Cc', 'Bcc', 'X-From', 'X-To', 'X-cc', 'X-bcc', 'X-Folder', 'X-Origin', 'X-FileName')
            temp_info = ['NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA', 'NA']
        
            # Iterate through each field and parse relevant information for each, if possible.
            for j_indx in range(len(desired_field_tuple)):
                if (j_indx==0)|(j_indx==len(desired_field_tuple)-1):
                    temp_info[j_indx] = parseInfo(contents, desired_field_tuple[j_indx], '')
                else:
                    for k in range(j_indx+1, len(field_tuple)):
                        if (j_indx>=6):
                            k = k+3
                        if(k>len(field_tuple)-1):
                            break
                        parse_result = parseInfo(contents, desired_field_tuple[j_indx], field_tuple[k])
                        if parse_result!='NA':
                            if parse_result != ' ':
                                temp_info[j_indx] = parse_result
                            break
            count = count + 1
            print('File(s) Read: ', count) # DEBUG PRINT
            
            #    For date_info, we know what each index corresponds to (based on the corresponding
            #    Regex from regExDict['DateParser']):
            #      [0]: Weekday
            #      [1]: Day
            #      [2]: Month
            #      [3]: Year
            #      [4]: Time
            #      [5]: Timezone (numeric)
            #      [6]: Timezone (acronym)
            date_info = re.findall(regExDict['DateParser'], temp_info[1])
            if date_info != []:
                date_info = date_info[0]
                date = monthDict[date_info[2]]+'/'+date_info[1]+'/'+date_info[3]
                time = date_info[4]
                timezone = date_info[6]
            else:
                date = 'NA'
                time = 'NA'
                timezone = 'NA'
            temp_info.pop(1)
            temp_info.insert(1, date)
            temp_info.insert(2, time)
            temp_info.insert(3, timezone)
        
            # Replace all commas in any field with '(COMMA)' in order to export/save to
            #    to a CSV file.
            for x_indx in range(len(temp_info)):
                temp_info[x_indx] = temp_info[x_indx].replace(',','(COMMA)')
        
            #print('Parsed Info: ')
            #for elem in range(len(temp_info)):
            #    print(final_field_tuple[elem],': ', temp_info[elem])
            email_writer.writerow(temp_info)
        f.close()
