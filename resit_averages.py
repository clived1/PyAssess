# Python code to take in resit CS input grids (3 columns per student)
# and calculate the new averages from the resits based on the standard rules
# (see Judith's email 02-Sep-2021)
#
# This program does not determine progression rules/status/new resits etc.
# i.e. these must be done by hand. However, the averages should be correct in most cases
# This was decided because the full progam would not be ready in time for this year
#
# The code has been checked against the 2018/19 grids (Y1+Y2) and agree with the year averages
# Note: cumulative averages for Y2 (Y1+Y2 average) has not been checked so be careful!
#
# NOTES/LIMITATIONS:
#
# 1. Only recalculates new averages and credits using original/resit/exclude mark depending on standard rules (see Judith's email 02-Sep-2021)
# 2. Currently cannot output merged cells for 2 columns under each unit - output is 1 column with code added to the mark e.g. "|30 | _R |" would be "30_R"
# 3. Does not take higher mark if student in danger of failing i.e. no additional compensation, but this should be fairly rare
# 4. Yr2 cumulative averages have not been checked and may be incorrect
#
# 
# MODIFICATION HISTORY:
#
# 03-Sep-2021  C. Dickinson   Made a start reading in data
# 04-Sep-2021  C. Dickinson   Read in data
# 05-Sep-2021  C. Dickinson   Calculate averages including resits
# 06-Sep-2021  C. Dickinson   Many bug fixes after checking against 2019/2020 data - works perfectly now
#                             Improved output formatting for easier reading
# 07-Sep-2021  C. Dickinson   Changed R1 to deferall so to use new mark and not capped (see Ivan's email 07-Sep-2021)

################################################################
# INPUTS
################################################################
# imports
#from typing import Any
import pandas as pd
import numpy as np
import sys
#import matplotlib.pyplot as plt
#from scipy.stats import linregress
#import itertools
#from decimal import Decimal
pd.options.mode.chained_assignment = None  # default='warn'

# Main input parameters
classyear   = 2                                       # Year 1 or 2
outdir      = './'                                     # Output directory ('./' for current directory)
sheet_name  = 'Resits'                                 # Excel sheet name (can set to anything)

# Year 1
if (classyear == 1):
    indir       = 'CS Resit exams grids/'                  # Directory containing input grid file (xls format)
    filename    = '1st year CS resit grid_06.09.21.xlsx'   # Excel filename for input grid
    outfilename = 'yr1_resits_averages.xlsx'                             # Excel filename for output grid

    #indir       = 'Resit exams grids 18-19/'
    #filename = '1styrresit_30_07_19_reformatted.xlsx'      # For checking (needs slight reformatting - 4 comment rows, Unit 1, Unit 2 etc.

# Year 2
if (classyear == 2):
    indir       = 'CS Resit exams grids/'                  # Directory containing input grid file (xls format)
    filename    = '2nd year CS resit grid_06.09.21.xlsx'   # Excel filename for input grid
    outfilename = 'yr2_resits_averages.xlsx'                             # Excel filename for output grid

    #indir       = 'Resit exams grids 18-19/'
    #filename    = '2ndyrresit_30_07_19_reformatted.xlsx'      # For checking
    
# combine directory and filename to be used later
filename = indir + filename
outfilename = outdir + outfilename

################################################################
# FUNCTIONS
################################################################
# used for skipping first few rows of CS grid (note system dependent)
def rowskiplogic(index):
    if index==0 or index==1 or index==2 or index==3:
       return True
    return False

def rowskiplogiccourse(index):
    if index==0 or index==1 or index==2 or index==3:
       return True
    return False

################################################################
# some data definitions
################################################################

# these are where units may have different credits for the marks vs progression (*make sure these are integers, not floats!)
credweightunits={'PHYS20040':10,  # main general paper (doesn't count towards progression/resits, but does count towards marks)
'PHYS20240':6,            # shorter version worth only 6 (M+P,Phys/Phil, 2nd/3rd year direct entry) 
'PHYS20811':5,            # Professional development ***CD: changed from 9 to 5 in AY2021
'PHYS20821':5,            # for the few students resitting the year this course still here
'PHYS30010':10,           # General paper (doesn't count towards progression/resits, but does count towards marks)
'PHYS30210':6,           # General paper (short version for M+P, Phys/Phil, 2nd/3rd year direct entry)
'PHYS30811':3}             # Added back for those few students re-sitting
#'PHYS20030':0,            # Peer-Assisted Study Sessions (PASS) - no marks, no credits, but here just in case 
#'ULGE21030':0,            # 
#'ULFR21030':0,
#'ULJA21020':0,
#'ULRU11010':0,
#'MATH35012':0,
#'COMP39112':0,
#'MATH49102':0}
                
# BELOW IS JUST FOR TESTING WITH 2018/19 DATA!! ***REMOVE/COMMENT OUT!!!
#credweightunits={'PHYS20040':10,
#'PHYS20240':6,  #6  
#'PHYS20811':9,  #9
#'PHYS20821':5,
#'PHYS30010':10,
#'PHYS30210':6,
#'PHYS30811':3,
#'PHYS20030':0,
#'ULGE21030':0,
#'ULFR21030':0,
#'ULJA21020':0,
#'ULRU11010':0,
#'MATH35012':0,
#'COMP39112':0,
#'MATH49102':0
#}

# courses to completley ignore because they don't have a mark e.g. tutorials, PASS etc. 
ignore_courses={'MPHYS',         # not a course
                'MPHYSON',       # not a course
                'MPHYSHON',      # not a course
                'PHYS10000',    # tutorials/similar
                'PHYS20000',    # tutorials/similar
                'PHYS30000',    # tutorials/similar
                'PHYS40000',    # tutorials/similar
                'PHYS10010',
                'PHYS10020',
                'PHYS10030',
                'PHYS10022',
                'PHYS11000',
                'PHYS21000',
                'PHYS31000',
                'PHYS41000',
                'PHYS20030',
                'PHYS19990',    # PASS Peer=Assissted Self-Study
                'PHYS29990',    # PASS Peer=Assissted Self-Study
                'PHYS39990',    # PASS Peer=Assissted Self-Study
                'PHYS49990',    # PASS Peer=Assissted Self-Study
                'MATH S100',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]
                'MATH S200',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]
                'MATH S300',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]
                'MATH S400',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]
                'MATH    S'     # sometimes this format
}


################################################################
# MAIN PROGRAM
################################################################


# open file and clean file
try:
    if (filename.split('.')[-1] == 'csv'): df = pd.read_csv(filename, skiprows=lambda x: +rowskiplogic(x), dtype='str',encoding = "ISO-8859-1")
    else: df=pd.read_excel(filename,skiprows=lambda x: +rowskiplogiccourse(x),dtype='str')
except:
    print('\nERROR reading the main input file. Please check filename and/or directory and make sure it is xls/xlsx/csv format...\n')
    sys.exit(0)
    
df = df.replace(np.nan, '', regex=True)  # Remove NaNs
df_orig = df.copy()  # keep original copy for analysisstudentfinalyear which contains extra codes in those columns
dftemp = df.loc[:, ~df.columns.str.contains('^Unnamed',na=False)]   # remove Unnamed columns
df_out  = dftemp.copy()  # for making changes on output later


columns=list(dftemp.columns.values)  # list of output column names (not including extra Unnamed ones)

dfallstudents= pd.DataFrame(columns=columns)  # data frame for all students

################################################################
# loop through each student and their marks, adding codes back from the 2nd column in the original file to the output
sids = df.loc[:,'Emplid']

# get IDs. select column of IDs and then every 3rd one using ::3
# Start at 1 because first row is blank (***fix later)
#sids=list(df.loc[:,'Emplid'][1::3])
sids=list(df.loc[:,'Emplid'])    # don't need to do ::3 because we can filter below
sids = [str for str in sids if str.isnumeric()]     # get just the IDs
nsids = len(sids)

# loop over IDs
counter = 0
for anid in sids:
    counter += 1
    print('Processing ID={0:s} ({1:d}/{2:d})'.format(anid,counter,nsids))

    selected=df.loc[df['Emplid'] == anid]
    #print(selected)
    
    # get index number of this ropw
    selectedindex=selected.index

    # extract course names (top row of target student in grid)
    studata=(df.iloc[selectedindex])
    studata=studata.iloc[0]
    
    # extract original course marks (2nd row in grid)
    stumarks=(df.iloc[selectedindex+1])
    stumarks=stumarks.iloc[0]

    # extra resits course marks (3rd row in grid)
    stumarks2=(df.iloc[selectedindex+2])
    stumarks2=stumarks2.iloc[0]
    
    # data frame for this student
    dfthisstudent = pd.DataFrame(columns=columns)

    # Check which column has Programme
    index = [idx for idx, s in enumerate(columns) if 'Plan' in s or 'Prog' in s][0]
    progtaken=studata[index]        # get prog     
    
    # variables for each student
    creditstaken=0
    creditstaken2=0
    creditspassed=0
    creditspassed2=0
    mean=0
    meancredsum=0
    meancredsum2=0
    creditsformean=0
    creditsformean2=0
    creditsformeansum=0
    creditsformeansum2=0
    creditsexcluded=0

    mathunit=0
    mathscreditstaken = 0
    mathscreditsmeansum = 0
    mathscreditspassed=0

    # maths sums. only relevant for maths/phys students (0=physics, 1=maths/physics student)
    if('Math' in progtaken): mathsphys=1
    else: mathsphys=0
    
    # loop over student courses
    for i in range(studata.size-1):

        # spot a valid course. (An upgrade would do this automatically)
        if( ('PHYS' in studata[i] or'EART'in studata[i] or 'HSTM'in studata[i]  or 'UCOL'in studata[i] or 'MATH'in studata[i]
            or 'BIOL'in studata[i] or 'PHIL'in studata[i] or 'UCIL'in studata[i] or 'ECON' in studata[i] 
            or 'COMP' in studata[i] or 'BMAN' in studata[i] or 'HSTM' in studata[i] or 'MCEL' in studata[i]
            or 'MACE' in studata[i] or 'ULBS' in studata[i] or 'ULCH' in studata[i] or 'ULIT' in studata[i]
            or 'ULPT' in studata[i] or 'ULGE' in studata[i] or 'ULFR' in studata[i] or 'ULJA' in studata[i]
            or 'ULSP' in studata[i] or   'ULAR' in studata[i] or 'ULBS' in studata[i] or 'ULCH' in studata[i]
            or 'ULDU' in studata[i] or 'ULFR' in studata[i] or 'ULGE' in studata[i] or 'ULHB' in studata[i]
            or 'ULIT' in studata[i] or 'ULKR' in studata[i] or 'ULJA' in studata[i] or 'ULPE' in studata[i]
            or 'ULPT' in studata[i] or 'ULRU' in studata[i] or 'ULSP' in studata[i] or 'ULTU' in studata[i]
            or 'ULUR' in studata[i] or 'MUSC' in studata[i] or 'ITAL' in studata[i] or 'FREN' in studata[i]
            or 'GERM' in studata[i] or 'POLI' in studata[i] or 'EEEN' in studata[i])
            ):
                # get course name
                coursename = studata[i][0:9]

                # ignore irrelevant courses with no marks/credits
                if(coursename.strip() in ignore_courses): continue
                
                # reset variables.
                # exclude is MC exclusion.
                excludethiscourse = 0
                excludethiscourse2 = 0
                mainmarkfound = 0
                thiscredit=0
                orig_resitmark = 0

                # add code back to stu marks 
                if (stumarks[i] != ''):
                    if (not stumarks[i][-1].isdigit()):
                        if (not stumarks[i][1].isdigit()): stumarks[i] = stumarks[i][0] + '_' + stumarks[i][1:]
                        elif (not stumarks[i][2].isdigit()): stumarks[i] = stumarks[i][0:2] + '_' + stumarks[i][2:] 
                        stumarks[i] = stumarks[i].replace('__','_')  # remove extra '__' if exists
                        
                    if (stumarks[i+1] != ''):
                        stumarks[i] = stumarks[i] + '_' + stumarks[i+1]
                        stumarks[i] = stumarks[i].replace('__','_') # remove extra '__' if exists
                    
                if (stumarks2[i] != ''):
                    if (not stumarks2[i][-1].isdigit()):
                        if (not stumarks2[i][1].isdigit()): stumarks2[i] = stumarks2[i][0] + '_' + stumarks2[i][1:]
                        elif (not stumarks2[i][2].isdigit()): stumarks2[i] = stumarks2[i][0:2] + '_' + stumarks[i][2:] 
                        stumarks2[i] = stumarks2[i].replace('__','_')  # remove extra '__' if exists
                        
                    if (stumarks2[i+1] != ''):
                        stumarks2[i] = stumarks2[i] + '_' + stumarks2[i+1]
                        stumarks2[i] = stumarks2[i].replace('__','_') # remove extra '__' if exists
                        
                # if a mark ends R or R in 2nd column, put R back
                #if (stumarks[i].find('R') >= 0):
                #    stumarks[i] = stumarks[i][:stumarks[i].find('R')] + '_' + stumarks[i][stumarks[i].find('R'):]
                #elif (stumarks[i+1].find('R') >= 0):
                #    stumarks[i] = stumarks[i] + '_' + stumarks[i+1][stumarks[i+1].find('R'):]
                    
                # if a mark end in C or C in 2nd column
                #if (stumarks[i].find('C') >= 0):
                #    stumarks[i] = stumarks[i][:stumarks[i].find('C')] + '_' + stumarks[i][stumarks[i].find('C'):]
                #elif (stumarks[i+1].find('C') >= 0):
                #    stumarks[i] = stumarks[i] + '_' + stumarks[i+1][stumarks[i+1].find('C'):]

                # if a mark end in X or X in 2nd column e.g. XL_D
                #if (stumarks[i].find('X') >= 0):
                #    stumarks[i] = stumarks[i][:stumarks[i].find('X')] + '_' + stumarks[i][stumarks[i].find('X'):]
                #elif (stumarks[i+1].find('X') >= 0):
                #    stumarks[i] = stumarks[i] + '_' + stumarks[i+1][stumarks[i+1].find('X'):]
                    
                # if a mark end in A or A in 2nd column
                #if (stumarks[i].find('A') >= 0):
                #    stumarks[i] = stumarks[i][:stumarks[i].find('A')] + '_' + stumarks[i][stumarks[i].find('A'):]
                #elif (stumarks[i+1].find('A') >= 0):
                #    stumarks[i] = stumarks[i] + '_' + stumarks[i+1][stumarks[i+1].find('A'):]
                    
                # If no mark, ignore, and just continue to next one
                if (stumarks[i] == ''):
                    print("*WARNING: no mark "+ coursename + ' ' + anid)
                    continue

                # get numerical mark (removing _R or _C if exists)
                if (stumarks[i].isdigit()): thismark = float(stumarks[i])
                else: thismark = float(stumarks[i][0:stumarks[i].find('_')])

                # Check to see if there is a resit mark and add back - if not, use original mark
                if (stumarks2[i] == ''):
                    thismark2 = thismark
                else:
                    if (stumarks2[i].find('R') >= 0):
                        stumarks2[i] = stumarks2[i][:stumarks2[i].find('R')] + '_' + stumarks2[i][stumarks2[i].find('R'):]
                    elif (stumarks2[i+1].find('R') >= 0):
                        stumarks2[i] = stumarks2[i] + '_' + stumarks2[i+1][stumarks2[i+1].find('R'):]
                    
                    # if a mark end in C or C in 2nd column
                    if (stumarks2[i].find('C') >= 0):
                        stumarks2[i] = stumarks2[i][:stumarks2[i].find('C')] + '_' + stumarks2[i][stumarks2[i].find('C'):]
                    elif (stumarks2[i+1].find('C') >= 0):
                        stumarks2[i] = stumarks2[i] + '_' + stumarks2[i+1][stumarks2[i+1].find('C'):]

                    # if a mark end in A or A in 2nd column
                    if (stumarks2[i].find('A') >= 0):
                        stumarks2[i] = stumarks2[i][:stumarks2[i].find('A')] + '_' + stumarks2[i][stumarks2[i].find('A'):]
                    elif (stumarks2[i+1].find('A') >= 0):
                        stumarks2[i] = stumarks2[i] + '_' + stumarks2[i+1][stumarks2[i+1].find('A'):]
                
                # get numerical mark for resit (removing _R or _C if exists)
                if (stumarks2[i] != ''):
                    if (stumarks2[i].isdigit()):
                        thismark2 = float(stumarks2[i])
                        orig_resitmark = thismark2
                    else:
                        thismark2 = float(stumarks2[i][0:stumarks2[i].find('_')])
                        orig_resitmark = thismark2
                        
                # Apply resit rules to see if original or resit mark is used or capped at 30%
                if (thismark2 == '' or stumarks2[i].find('XN') >= 0):
                    thismark2 = thismark
                if (stumarks[i].find('R') >= 0 and thismark >= 29.95): # should work for 18/19 still
                    thismark2 = thismark
                if (stumarks[i].find('R') >= 0 and thismark < 29.95):  # should work for 18/19 still
                    if (thismark2 > 30): thismark2 = 30
                if (stumarks[i].find('R1') >=0 ):  # R1 means deferall and should not be kept (should work for 18/19 still because no R1 used then)
                    thismark2 = orig_resitmark
                    
                # get credits

                # must be a straightforward mark or a progression
                # dig out credits taken
                lbindex = studata[i].find('(')
                rbindex = studata[i].find(')')
                # print(lbindex, rbindex)
                # print(studata[i][lbindex+1:rbindex])
                thiscredits = int(studata[i][lbindex + 1:rbindex])  # should be integer
                #print('using main mark credits ',thiscredits)
                #print(targetstudent,studata[i][0:9],stumarks[i],stumarks[i+1],thiscredits)

                # credits for this course
                creditsformean = thiscredits

                #print('creditsformean ',creditsformean)
                # if zero credits, may need credits for credit weighted mean
                if (creditsformean == 0): 
                    try:
                        creditsformean = credweightunits[coursename]
                    except: print('*WARNING: I do not know the credit weight for the mean of ', coursename)
                    
                #print('Mark = ',thismark, ' with ', creditsformean, ' credits for ',coursename)

                # resits
                thiscredits2 = thiscredits
                creditsformean2 = creditsformean
                        
                # panic if credits not known!
                if (creditsformean == 0):
                    print('*WARNING: I do not know the credit weight for the mean of ', coursename)

                # Exlcude mark from year mark if X/XL/X1 etc. except XN
                if (stumarks[i].find('X') >= 0 and stumarks[i].find('N') < 0):
                    #print("exclude!")
                    excludethiscourse = 1
                    creditsexcluded+=creditsformean
                                    
                if (stumarks[i].find('XN') >= 0):
                    #print("missed with no reason. NOT exclude!")
                    excludethiscourse = 0

                # exclude resit if 'X' in original mark and no resit mark
                if (stumarks[i][-1] == 'X' and stumarks2[i] == ''):
                    excludethiscourse2 = 1

                # exclude resit if has X but not XN
                if (stumarks2[i].find('X') >= 0 and stumarks2[i].find('N') < 0):
                    excludethiscourse2 = 1
                    
            #####################
            
                # log the credits taken by the student, first for physics, then maths (if maths/phys)
                creditstaken += thiscredits
                creditstaken2 += thiscredits2
                
                if (not excludethiscourse): creditsformeansum += creditsformean

                # if a maths course for recording M+P separately (for M+P students)
                if ('MATH' in studata[i]):  mathunit = 1

                if (mathunit == 1 and not excludethiscourse):
                    mathscreditstaken += thiscredits
                    
                # update meancredsum (used for mean)
                if(not excludethiscourse):
                    meancredsum+=thismark*creditsformean               
                    if(mathunit==1):
                        mathscreditsmeansum+=thismark*creditsformean

                # resits if not excluded
                if (not excludethiscourse2):
                    creditsformeansum2 += creditsformean2
                    meancredsum2 += thismark2*creditsformean2

                # credits passed
                if (thismark >= 39.95): creditspassed += thiscredits
                if thismark2 >= 39.95: creditspassed2 += thiscredits2  # 
                elif (orig_resitmark >= 39.95): creditspassed2 += thiscredits2 # for resits passed but capped/use original mark
                                        
                # TESTING ONLY
                print(anid,coursename, thismark, thismark2, excludethiscourse2, stumarks[i], stumarks2[i], orig_resitmark)
                    
    # TESTING ONLY
    #print('*',meancredsum2, creditsformeansum2)
    
    # compute mean credit sum
    if(creditsformeansum>0):
        meancredsum=float(meancredsum)/creditsformeansum
        #print('Mean mark calc:',stuname, meancredsum,creditsformeansum,creditstaken)
    else:
        meancredsum=0

    if(mathscreditstaken>0):
        mathscreditsmeansum=float(mathscreditsmeansum)/mathscreditstaken
    else:
        mathscreditsmeansum=0

    # resit average
    meancredsum2 = float(meancredsum2)/creditsformeansum2
        
    # 1.d.p.
    meancredsum = round(meancredsum+0.0000001,1)
    meancredsum2 = round(meancredsum2+0.0000001,1)
        
    # log student units/marks to data frame 
    dfthisstudent= dfthisstudent.append(studata)
    dfthisstudent = dfthisstudent.append(stumarks)
    dfthisstudent = dfthisstudent.append(stumarks2)

    # log data we have made in new columns in data frame
    dfthisstudent['Units\nTaken'].iloc[0]=creditstaken
    dfthisstudent['Units\nPassed'].iloc[0] = creditspassed
    dfthisstudent['Units\nTaken'].iloc[1]=creditstaken2
    dfthisstudent['Units\nPassed'].iloc[1] = creditspassed2

    # put in averages
    if (classyear == 1):
        dfthisstudent['Yr Mk\n/ GPA'] = [meancredsum,meancredsum2,'']
        #dfthisstudent['Yr Mk\n/ GPA'].iloc[1] = meancredsum2

    if (classyear == 2):
        dfthisstudent['Yr Mk\n/ GPA'] = [meancredsum, meancredsum2,'']
        #dfthisstudent['Yr Mk\n/ GPA'].iloc[1] = meancredsum2
        
    # If Y2, calculate overall mark if previous mark available
    if (classyear == 2 and 'Pr' in ''.join(columns)):

        yr1mark = float(dfthisstudent['Pr. Yr Mk\n/ GPA'].values[0])
        yr1mark = round(yr1mark+0.000001,1)
        
        if( ('BSc' in progtaken) and not ('Math' in progtaken)):
            overallmark = 0.25*yr1mark + 0.75*meancredsum
            overallmark2 = 0.25*yr1mark + 0.75*meancredsum2
            
        if( ('MPhys' in progtaken) and not ('Math' in progtaken)):
            overallmark = (0.06/0.25)*yr1mark + (0.19/0.25)*meancredsum
            overallmark2 = (0.06/0.25)*yr1mark + (0.19/0.25)*meancredsum2
            
        if( ('BSc' in progtaken) and ('Math' in progtaken)):
            overallmark = 0.25*yr1mark + 0.75*meancredsum
            overallmark2 = 0.25*yr1mark + 0.75*meancredsum2

        if( ('MPhys' in progtaken) and ('Math' in progtaken)):
            overallmark = (0.06/0.25)*yr1mark + (0.19/0.25)*meancredsum
            overallmark2 = (0.06/0.25)*yr1mark + (0.19/0.25)*meancredsum2
            
        if ( ('MPhys' in progtaken and not('Math' in progtaken) and 'Study' in progtaken)):
            overallmark = (0.08/0.31)*yr1mark + (0.23/0.31)*meancredsum
            overallmark2 = (0.08/0.31)*yr1mark + (0.23/0.31)*meancredsum2

        overallmark = round(overallmark+0.000001,1)
        overallmark2 = round(overallmark2+0.000001,1)

        dfthisstudent['Cumul.\nGPA'].iloc[0] = overallmark
        dfthisstudent['Cumul.\nGPA'].iloc[1] = overallmark2

        dfthisstudent['Pr. Yr Mk\n/ GPA'].iloc[1] = dfthisstudent['Pr. Yr Mk\n/ GPA'].iloc[0]
        
    # Finally append to main dataframe
    dfallstudents = dfallstudents.append(dfthisstudent)

    #sys.exit(0)          # TESTING ONLY
    #if (anid == '10443485'): sys.exit(0)
    
################################################################           
df_out = dfallstudents.copy()

# If no average columns add them
if (not 'Yr Mk' in ''.join(columns)): columns.append('Yr Mk\n/ GPA')


################################################################           
# output to Excel spreadsheet (strings_to_numbers option so not to store unit numbers as text)
try:
    writer = pd.ExcelWriter(outfilename, engine='xlsxwriter',options={'strings_to_numbers': True})
except:
    print('\nERROR writing out file. Please check directory...\n')
    sys.exit(0)

# output columns to excel spredsheet directly from Data Frame
df_out.to_excel(writer,index=False,sheet_name=sheet_name,columns=columns)

# Change column width/formatting before finally writing out
workbook = writer.book
worksheet = writer.sheets[sheet_name]
i = 0
for column_str in columns: # loop over each column
    col_idx = i  # df_out.columns.get_loc(column_str)
    # set the widths
    if (column_str == 'Emplid'): col_width=9
    elif (column_str == 'Name'): col_width=15
    elif (column_str == 'PSI'): col_width=4
    elif (column_str == 'Plan'): col_width=20
    elif (column_str.find('Admit') >=0): col_width=8
    elif (column_str.find('Unit') >= 0): col_width=12
    elif (column_str.find('Units') >= 0): col_width=10
    elif (column_str.find('AS') >= 0): col_width=8
    elif (column_str.find('Pr.') >= 0): col_width=12
    elif (column_str.find('Yr Mk') >= 0): col_width=12      
    else: col_width=12  # default column width
    writer.sheets[sheet_name].set_column(col_idx, col_idx, col_width)
    i += 1
    
# make banded rows and slightly larger cells for easier viewing
format1_grey = workbook.add_format({'bg_color': '#E0E0E0'})

cellheight = 15  # (Excel default is 15)
nrows = df_out.count()[0]
# every 3 are banded/non-banded
for row in range(1, nrows, 6):
    worksheet.set_row(row, cellheight, cell_format=format1_grey)
    worksheet.set_row(row+1, cellheight, cell_format=format1_grey)
    worksheet.set_row(row+2, cellheight, cell_format=format1_grey)
    worksheet.set_row(row+3, cellheight)
    worksheet.set_row(row+4, cellheight)
    worksheet.set_row(row+5, cellheight)
    
# Write out and finish
writer.save()

# Final output before finishing
print('\nProcessing complete and output written to {:s}'.format(outfilename))    
print('\nNote: status and progression rules have not been applied - only averages and credits have been updated\n')
