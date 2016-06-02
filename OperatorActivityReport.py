import re   #String comparison library
import csv  #CSV output library
import glob #Multiple file input library
import os.path #File checking library

###
# Operator/Branch Activity Parser
# Created by: Christopher Campbell (Ccampbell@p1fcu.org)
#
# This program navigates to a folder, reads through every file in the folder and parses the data into a csv file
# To use: Change path and outfile strings to your desired input path and output file and run this script
# Input: Folder containing ACCT.OPER.STATS files (~line 47)
# Output: .csv file with the parsed data (~line 48)
###

###
# get_branch_keys(infile, BranchKey)
# Reads through every branch and adds it to a list in order, then we can just increment a counter to find the associated branch name
# Parameters: infile - path to file being read
#             BranchKey - list of branches in order of keys assigned, output
# Returns: BranchKey
###
def get_branch_keys(infile, BranchKey):
    for scanline in infile: #For every line in the file we open
        if re.search(r'-----', scanline):   #If line contains '-----'(Always the line before a branch
            x = infile.readline()           #Read to next line to skip the '-----' line
            x = x.split('  ')               #Using 2 spaces as a delimeter break line into chunks
            if(len(x) > 2):                 #If the line has something where branch name should be
                br = x[1].strip()           #Strip whitespace from around branch name
                BranchKey.append(br)        #Add branch name to our list
    return BranchKey                        #Return the list

###
# header(writer)
# Writes a header on the top of the file with column names
# Parameters: writer - CSV Writer to print output to a .csv file
# Returns: none
def header(writer):
    writer.writerow(["Operator #", "Operator", "Branch", "Before 0700", "700-800", "800-900", "900-1000", "1000-1100", "1100-1200", "1200-1300", "1300-1400", "1400-1500", "1500-1600", "1600-1700", "1700-1800", "1800-1900", "1900-2000", "After 2000", "Grand Total"])

###
# parse_operator_activity()
# Main function opens input files, extracts desired fields and prints them to an output file
# Parameters: none
# Returns:none
###
def parse_operator_activity():
    ###
    # The following two lines allow hardcoding the path and outfile destination
    # To use them comment out the lines up until but not including writer = csvwriter(line ~ 74)
    ###
    # path = r"X:\Python Scripts\OperatorActivityReport\Input\*"        #Path to folder containing our input files
    # outfile = open(r"X:\Python Scripts\OperatorActivityReport\Output.csv", "w+", newline='')  #Path to our output file
    path = ' '                  #Set path and outfile to empty strings to initialize them
    outfile = ' '
    # Collect input and output location from user
    while path == ' ':                              #While path is an empty string prompt for input
        pathentry = input('Enter an input location:  ')
        if os.path.exists(pathentry):                       #If path is a valid path
            path = pathentry                                #Sets a test variable
            print("Valid path")
        else:
            pathentry = input('Please Enter a valid input location such as: X:\input\ ')
    while outfile == ' ':                           #While outfile location is empty get input
        outfileentry = input('Enter an output destination and file: ')
        if os.path.isfile(outfileentry):
            print('There is already a file of that name in existence, please choose another ')
        else:
            try:
                outfile = open(outfileentry, "w+", newline='')          #Open the appropriate outfile in write mode
                print("Valid output file")
            except OSError:
                print("Invalid output destination")
                outfile = open("C:\Public\Public Documents", "w+", newline = '')

    #If hardcoding paths, comment until here
    writer = csv.writer(outfile, delimiter=',', quotechar='|', quoting=csv.QUOTE_MINIMAL) #CSV writer assignment
    global firstfile                                            #global flag for printing header
    firstfile = 1                                               #Assign flag to 1 to start
    file_collector = path + "*"
    infilelist = glob.glob(file_collector)                                #Create a list of everything in input path
    for file in infilelist:                                     #For every file in our input path
        infile = open(file, "r+")                           #Open the file
        infile.seek(1)                                      #Skip to second line
        first_line = infile.readline()                     #Chomp BOM characters
        title_line = infile.readline()                      #Read title line
        date = title_line[113:121]                          #Pull date from title line
        if(firstfile == 1):                                 #If this is the first file print header and switch off flag
            header(writer)
            firstfile = 0
        BranchKey = []                                  #Initialize BranchKey as a list
        get_branch_keys(infile, BranchKey)              #Call function to populate BranchKey list
        count = 0                                       #Count variable for use in BranchKey pairing
        infile.seek(1)                                  #Skip to line 2
        global branchflag                                     #Initialize flag for reseting count on branch
        branchflag = False

        for line in infile:                             #For each line in current file
            if re.match(r'POTLATCH', line) or re.match(r'\f', line):    #If line starts with POTLATCH or Form Feed
                #print ("Got POTLATCH or FF")
                continue                                                #Skip it
            elif re.match(r'GRAND', line):              #If line is GRAND TOTAL, print it without false fields
                operator = line[0:26].strip()           #Assign variables by parsing line and removing whitespace
                field1 = line[28:32].strip()
                field2 = line[34:38].strip()
                field3 = line[40:44].strip()
                field4 = line[46:50].strip()
                field5 = line[52:56].strip()
                field6 = line[58:62].strip()
                field7 = line[64:68].strip()
                field8 = line[70:74].strip()
                field9 = line[76:80].strip()
                field10 = line[82:86].strip()
                field11 = line[88:92].strip()
                field12 = line[94:98].strip()
                field13 = line[100:104].strip()
                field14 = line[106:110].strip()
                field15 = line[112:116].strip()
                GrandTotal = line[118:124].strip()
                #Following line is for debugging
                #print(' ' + ',' + operator + ',' + ' ' + ',' + field1 + ',' + field2 + ',' + field3 + ',' + field4 + ',' + field5 + ',' + field6 + ',' + field7 + ',' + field8 + ',' + field9 + ',' + field10 + ',' + field11 + ',' + field12 + ',' + field13 + ',' + field14 + ',' + field15 + ',' + GrandTotal + ',' + date)
                writer.writerow([operator_number, operator, branch, field1, field2, field3, field4, field5, field6, field7, field8, field9, field10, field11, field12, field13, field14, field15, GrandTotal, date])    #Writes our values to the output file
            elif re.match(r'\w', line) or re.match(r'\d', line):    #If line starts with a letter or number its an operator
                branch = BranchKey[count]                       #Assign proper branch name
                operator_number = line[0:3].strip()             #Assign values by parsing line and stripping whitespace
                operator = line[4:26].strip()
                field1 = line[28:32].strip()
                field2 = line[34:38].strip()
                field3 = line[40:44].strip()
                field4 = line[46:50].strip()
                field5 = line[52:56].strip()
                field6 = line[58:62].strip()
                field7 = line[64:68].strip()
                field8 = line[70:74].strip()
                field9 = line[76:80].strip()
                field10 = line[82:86].strip()
                field11 = line[88:92].strip()
                field12 = line[94:98].strip()
                field13 = line[100:104].strip()
                field14 = line[106:110].strip()
                field15 = line[112:116].strip()
                GrandTotal = line[118:124].strip()
                #Debugging print line
                #print(operator_number + ',' + operator + ',' + branch + ',' + field1 + ',' + field2 + ',' + field3 + ',' + field4 + ',' + field5 + ',' + field6 + ',' + field7 + ',' + field8 + ',' + field9 + ',' + field10 + ',' + field11 + ',' + field12 + ',' + field13 + ',' + field14 + ',' + field15 + ',' + GrandTotal + ',' + date)
                writer.writerow([operator_number, operator, branch, field1, field2, field3, field4, field5, field6, field7, field8, field9, field10, field11, field12, field13, field14, field15, GrandTotal, date])    #Writes values to output file
                if branchflag is True:          #If this line is a new branch increment counter so following operators are in next branch
                    branchflag = False
                    count = count + 1
                else:
                    continue


            elif re.search(r'-----', line):        #If this line is the '-----' line the next line is a new branch, set the flag
                branchflag = True


parse_operator_activity()  #Call to run our function, this is what runs the program

