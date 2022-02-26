# Physics and Astronomy Exam assessment code# Takes in exam grids (csv/xls/xlsx files) and produces output exam grids, with averages, degree classes etc.# This version is for AY2021/2022 and is a complete re-write of the previous version (PyAssess2021.py)# that should, in principle, work for any years.## This script requires Pandas with xlswriter (default writer for pandas, as long as v>0.16)## Checked on ????## Known issues to keep in mind:## 1. Cannot handle "PRO" marks correctly. At the moment, it just ignores this code so the mark will be incorrect!!#    N.B. they should be in the donotprocess list to ensure they are dealt with separately)# 2. Does not not fully handle Physics with Study in Europe 2-3 progression (needs S3 mark which is not available/easy to compute).#    N.B. The code treats them as normal MPhys students - the Phys/Euro coordinator will check these few by hand.# 3. Does not handle M+P aboard students with 3A or A in the programme name# # FUTURE To-do list:## 1. Make module to allow for different year rules e.g. AY=1819, AY=2021 etc.# 2. Put in specific rules for 4th year MPhys fails and determine BSc grade (p78 of handbook)# 3. Tidy up progression and try to fix any loop holes still remaining# 4. Automatically work out how many info lines are at top in case it is not standard 4.## MODIFICATION HISTORY## 23-Feb-2022  C. Dickinson    Start from scratch #################################################################### importsimport pandas as pdimport numpy as npimport sys#from typing import Any#import matplotlib.pyplot as plt#import itertools#from decimal import Decimalpd.options.mode.chained_assignment = None  # default='warn'################################################################# FUNCTIONS################################################################# used for skipping first few rows of CS grid (note system dependent)def rowskiplogic(index):    if index==0 or index==1 or index==2 or index==3:       return True    return False################################################################# read in student data (usually from a csv file but can be excel)def read_data(df):    try:        if (filename.split('.')[-1] == 'csv'): df = pd.read_csv(filename, skiprows=lambda x: +rowskiplogic(x), dtype='str',encoding = "ISO-8859-1")        else: df=pd.read_excel(filename,skiprows=lambda x: +rowskiplogiccourse(x),dtype='str')    except:        print('\nERROR reading the main input file. Please check filename and/or directory and make sure it is xls/xlsx/csv format...\n')        sys.exit(0)    # clean up dataframe of NaNs but leave Unnamed columns for now       df = df.replace(np.nan, '', regex=True)  # Remove NaNs    return df################################################################# Remove Unnamed ('^Unnamed') columns from the dataframedef remove_unnamed_columns(df):        return df.loc[:, ~df.columns.str.contains('^Unnamed')]   # remove Unnamed columns################################################################# extract relevant students from df to another df by typedef extract_students_by_type(df,classyear,studtype):    # get indices of relevant students    if (classyear == 1 or classyear == 2 or classyear == 4):        if (studtype == 1):            idx = df.index[(df['Plan'].str.contains('MPhys')) | (df['Plan'].str.contains('BSc')) & (~df['Plan'].str.contains('Math'))]        else:            idx = df.index[(df['Plan'].str.contains('Math'))]                if(classyear==31):        if (studtype == 1):            idx = df.index[(df['Plan'].str.contains('MPhys'))]        else:            idx = df.index[(df['Plan'].str.contains('MMath'))]    #  Select BScs without maths    if(classyear==32):        if(studtype==1):            idx = df.index[(df['Plan'].str.contains('BSc')) & (df['Plan'].str.contains('Physics'))]        else:            df = df.index[(df['Plan'].str.contains('BSc')) & (df['Plan'].str.contains('Math'))]    # add indices for the 2nd row of each student record    idx2 = []   # array of full indices    idx = idx.to_list()  # array of the first row index    for i in range(len(idx)):        idx2.append(idx[i])        idx2.append(idx[i]+1)    dfthisstudent = df.iloc[idx2]    # extract all the relevant rows          return dfthisstudent################################################################# get all student ids from a dataframedef get_sids(df):    #sidsname=list(df.columns)[0]       #sids=(df.loc[:,sidsname])    sids = df.iloc[:,0]  # assumes first column is SID        if (sids.size < 1):        print('No student IDs found...check file and column for SIDs')        sys.exit(0)    sids = sids.to_list()  # make a list    sids = list(filter(None,sids))    # remove empty strings    return sids################################################################# get df for a given student ID from the main dfdef get_student(df,sid):        dfthisstudent=df.loc[df['Emplid'] == sid]  # select row with studentID= target    index=dfthisstudent.index.to_list()          # get index number of this row        dfthisstudent1 = df.loc[df.index == index[0]]    dfthisstudent2 = df.loc[df.index == index[0]+1]        dfthisstudent = pd.concat([dfthisstudent1, dfthisstudent2])    return dfthisstudent################################################################# convert degree class to a string for output griddef degclass_to_string(progtaken,degclass):    # deg classification    degclass_dict = {        4:"1",        3:"2:1",        2:"2:2",        1:"3",        0:"Ord.",        -1:"Fail",        -2:"NOT SET!!!"}    # degree type    degstr = progtaken.split('(')[0] + ' '                return degstr + degclass_dict[degclass]################################################################# remove unused columns (e.g. Admit Term, PSI)def remove_unused_columns(df):    unused_columns = ['Admit Term', 'PSI']   # beginning string of unnecessary columns    df.drop(unused_columns, inplace=True, axis=1)        return df################################################################# get course marks, credits and codesdef get_course_marks(df):    course_data = df.filter(regex='^Unit\ ',axis=1).values    # get unit information    # remove irrelevant courses            return course_data################################################################# function to resequence unit numbers, removing unused units and Unnamed columns, and putting codes into the 3rd rowdef resequence_units(df,ignore_list):    # add extra blank row    columns = df.filter(regex='^Unit\ ',axis=1).columns.to_list()   # get columns    blank_row = df.iloc[0,:].copy()   # make 3rd row (using copy not assigning to original df!)        blank_row[:] = ''   # make all elements empty string    df = df.append(blank_row, ignore_index=True)        # keep only useful units    unit_number = 0   # counter for (new) unit number    for col in columns:        coursename_full = df[col].values[0]  # including credits        coursename = df[col].values[0][0:9]  # coursename (first 9 characters)        coursemark = df[col].values[1]        idx = df.columns.get_loc(col)        coursecode = df.iloc[1,idx+1]        # if course to be ignored or is empty, drop the columm        if (coursename in ignore_list or coursename == ''):            df = df.drop(col, 1)                    #print('*', col, coursename, coursemark, coursecode)        if (not coursename in ignore_list and coursename != ''):  # if a useful unit include            unit_number += 1  # get unit number just in case not in numerical order            target_unit = 'Unit ' + str(unit_number)            #print(target_unit, col, coursename, coursemark, coursecode)            df[target_unit] = [coursename_full, coursemark, coursecode]                return df################################################################# function to get credits for each unit and credit weights for calculating averagesdef get_credits(df,credweights):    # add 2 extra rows for the credits/credit weights    columns = df.filter(regex='^Unit\ ',axis=1).columns.to_list()   # get columns    blank_row = df.iloc[0,:].copy()   # make 3rd row (using copy not assigning to original df!)        blank_row[:] = ''   # make all elements empty string    df = df.append(blank_row, ignore_index=True)  # for credits    df = df.append(blank_row, ignore_index=True)  # for credit weights    for col in columns:        coursename = df[col].values[0]        thiscredit = coursename[10:]        coursename = coursename[0:9]        print(coursename, thiscredit, len(thiscredit))        lbindex = thiscredit.find('(')        rbindex = thiscredit.find(')')        #print(lbindex, rbindex, thiscredit[lbindex+1:rbindex])        thiscredit = int(float(thiscredit[lbindex + 1:rbindex]))  # should be integer        df[col][3] = thiscredit   # credits as per the unit        # get credit weights of the averages        if (thiscredit == 0):  # if 0, use appropriate credit weight for average            try:                thiscredweight = credweightunits[coursename]            except: print('*WARNING: I do not know the credit weight for ', coursename)        else:            thiscredweight = thiscredit        df[col][4] = thiscredweight            return df################################################################################################################################# INPUTS AND READIING DATA################################################################# Define year (1=1st year, 2=2nd year, 31=3rd prog, 32=3rd complete ,4=4th year)classyear=1# Student type (1=physics, 2=Maths+Physics)#########studtype=2# Input directory for files (default is './' for current directory)indir = './'# Late penalty filename - put these in to store pre-penalty scores with _P codes# _P1 = premark >40, _P2 = premark =30-30, _P3 = premark < 30# If set to blank ('') then nothing is applid# File should be for both semesters, with col1=coursename, col2=ID, col3=pre-penalty mark, col4=post-penalty mark#late_penalty_filename = ''#late_penalty_filename = './data2021/Late-Penalties-both-semesters-JMcG-.xlsx'# Input and output files for each cohort# Input filename (filename) can be .csv or .xls/.xlsx - the code will automatically read it in whichever the format if(classyear==1):    filename = './data1819/1styr_28_06_19_postmcc.csv'    #filename = './data2021/1st year exam grid_10.06.21.csv'    #filename = './data2021/1st year exam grid_02.07.21_mcc anonymous.xlsx'        # output filename    if (studtype==1): outfilename = '1styear_Physics.AY2018.prelimfinal.xlsx'   # Physics filename    else: outfilename = '1styear_MathsPhysics.AY2018.prelimfinal.xlsx'   # M+P filename    elif (classyear == 2):    #filename = '/data1819/2ndyr_26_06_19.csv'    #filename = './data2021/2nd year exam grid_25.06.21.csv'    #filename = './data2021/2nd year exam grid_02.07.21_mcc anonymous.xlsx'    filename = './data2021/2nd year exam grid_16.07.21.xlsx'    # If CFfilename to be used, set deganalysis to 1 (below)    CFfilename = './data2021/2nd year carry forward.xlsx'        # output filename    if (studtype==1): outfilename = '2ndyear_Physics.AY2021.prelimfinal.xlsx'   # Physics filename    else: outfilename = '2ndyear_MathsPhysics.AY2021.prelimfinal.xlsx'   # M+P filenameelif(classyear==32):    #filename = './data1819/3rdyr_18_06_19_external.csv'    #filename = './data2021/3rd year exam grid_10.06.21.csv'    #filename = './data2021/3rd year exam grid_07.07.21_mcc anonymous.xlsx'    filename = './data2021/3rd year exam grid_10.07.21_postmcc.xlsx'        #CFfilename = './data1819/3rdyr_carryforward.csv'    CFfilename = './data2021/3rd year carry forward.xlsx'    # output filename    if (studtype==1): outfilename = 'FinalYear_BSc_Physics.AY2021.prelimfinal.xlsx'   # Physics filename    else: outfilename = 'FinalYear_BSc_MathsPhysics.AY2021.prelimfinal.xlsx'   # M+P filename    elif (classyear == 31):    #filename = '/data1819/3rdyr_18_06_19_external.csv'    #filename = './data2021/3rd year exam grid_10.06.21.csv'    #filename = './data2021/3rd year exam grid_07.07.21_mcc anonymous.xlsx'    filename = './data2021/3rd year exam grid_10.07.21_postmcc.xlsx'        #CFfilename = './data1819//3rdyr_carryforward.csv'    CFfilename = './data2021/3rd year carry forward.xlsx'    # output filename    if (studtype==1): outfilename = '3rdyear_MPhys.AY2021.prelimfinal.xlsx'   # Physics filename    else: outfilename = '3rdyear_MMath.AY2021.prelimfinal.xlsx'   # M+P filename    elif(classyear==4):    #filename = '/data1819/4thyr_18_06_19_external.csv'    #filename = './data2021/4th year exam grid_10.06.21.csv'    #filename = './data2021/4th year exam grid_07.07.21_mcc anonymous.xlsx'    filename = './data2021/4th year exam grid_10.07.21_with year abroad flags.xlsx'        #CFfilename='./data1819/4thyr_carryforward.csv'    CFfilename = './data2021/4th year carry forward.xlsx'    # output filename    if (studtype==1): outfilename = 'FinalYear_MPhys.AY2021.prelimfinal.xlsx'   # Physics filename    else: outfilename = 'FinalYear_MMath.AY2021.prelimfinal.xlsx'   # M+P filenameelse:    print('*ERROR: Classyear not defined correctly (should be 1, 2, 31, 32, or 4)')    sys.exit(0)    # final filenamefilename= indir + filename################################################################# some data definitions################################################################# Credits requiredcreditstogetMPHYS = 80  creditstogetBScgood=80creditstogetBSclower=60creditstogetMPHYSalgA = 75 # Used for algA creditstogetalgB = 70  # Used for algB# testing only for 2018/19 data! ***REMOVE/COMMENT OUT!!!creditstogetMPHYS = 80  creditstogetBScgood=80creditstogetBSclower=60creditstogetMPHYSalgA = 80 # Used for algA creditstogetalgB = 70  # Used for algB# boundaries for degree class (2 d.ps because marks are stored to 1 d.p.)boundaryfirst=69.95boundaryupper2=59.95boundarylower2=49.95boundarythird=39.95# borderlines for promotion consideration borderfirst = boundaryfirst - 3.0borderupper2 = boundaryupper2 - 3.0borderlower2 = boundarylower2 - 3.0borderthird = boundarythird - 4.0# any students to skip# For 2021 to omit exception students (see Suzanne's email 11-Jun-2021 and Y2 issues xls file and Judith email 08-Jul-2021) donotprocess={'10304702','10301241','10341954','9954785','9976148','9914290'}   # IDs should be strings! #donotprocess={}  # FOR TESTING ONLY!# define core for purpose of triggered resits i.e.  what gets resat if a student is going to have resits anyway.# This one is for studtype=1, for most Physics studentsisphysicscore={'PHYS10071','PHYS10101','PHYS10121','PHYS10191','PHYS10302','PHYS10342','PHYS10352','PHYS10372','PHYS20101','PHYS20141','PHYS20171','PHYS20252','PHYS20312','PHYS20352'}# M+P students have a different list in *addition* to the iscore list above# (only need 2nd year courses because all MATHs courses in Y1 must be passed - these courses add later when they are known)ismathscore={'MATH20401', 'MATH29142'}# Set the core list depending on whether M+P student or notif (studtype == 1): iscore = isphysicscoreelse: iscore = isphysicscore.union(ismathscore)#define what must be passed e.g. lab, BSc dissertation.mustpass={'PHYS10180','PHYS10280',          'PHYS20180','PHYS20280',          'PHYS30180','PHYS30280','PHYS30880',          'PHYS40181','PHYS40182'}# these are where units may have different credits for the marks vs progression (*make sure these are integers, not floats!)credweights={'PHYS20040':10,  # main general paper (doesn't count towards progression/resits, but does count towards marks)'PHYS20240':6,            # shorter version worth only 6 (M+P,Phys/Phil, 2nd/3rd year direct entry) 'PHYS20811':5,            # Professional development CD: changed from 9 to 5 in AY2021'PHYS20821':5,            # for the few students resitting the year this course still here'PHYS30010':10,           # General paper (doesn't count towards progression/resits, but does count towards marks)'PHYS30210':6,           # General paper (short version for M+P, Phys/Phil, 2nd/3rd year direct entry)'PHYS30811':3}             # Added back for those few students re-sitting#'PHYS20030':0,            # Peer-Assisted Study Sessions (PASS) - no marks, no credits, but here just in case #'ULGE21030':0,            # #'ULFR21030':0,#'ULJA21020':0,#'ULRU11010':0,#'MATH35012':0,#'COMP39112':0,#'MATH49102':0}noresitlist={}  # any units that can't be resit (other than 0 credit units (which are not resitable) like lab etc.)# courses to completley ignore because they don't have a mark e.g. tutorials, PASS etc. ignore_courses={'MPHYS',         # not a course                'MPHYSON',       # not a course                'PHYS10000',    # tutorials/similar                'PHYS20000',    # tutorials/similar                'PHYS30000',    # tutorials/similar                'PHYS40000',    # tutorials/similar                'PHYS10010',                'PHYS10020',                'PHYS10030',                'PHYS10022',                'PHYS11000',                'PHYS21000',                'PHYS31000',                'PHYS41000',                'PHYS20030',                'PHYS19990',    # PASS Peer=Assissted Self-Study                'PHYS29990',    # PASS Peer=Assissted Self-Study                'PHYS39990',    # PASS Peer=Assissted Self-Study                'PHYS49990',    # PASS Peer=Assissted Self-Study                'MATH S100',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]                'MATH S200',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]                'MATH S300',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]                'MATH S400',    # Maths Study module (0 credited) [last 0 omitted on prurpose because of extra gap]                'MATH    S'     # sometimes this format}                # BELOW IS JUST FOR TESTING WITH 2018/19 DATA!! ***REMOVE/COMMENT OUT!!!#credweightunits={'PHYS20040':10,#'PHYS20240':6,  #6  #'PHYS20811':9,  #9#'PHYS20821':5,#'PHYS30010':10,#'PHYS30210':6,#'PHYS30811':3,#'PHYS20030':0,#'ULGE21030':0,#'ULFR21030':0,#'ULJA21020':0,#'ULRU11010':0,#'MATH35012':0,#'COMP39112':0,#'MATH49102':0#}    ############################################################################################################################################################################################# NOW THE RUNNING PART OF THE CODE############################################################################################################################################################################################# get datadfallstudents = read_data(filename)     # read in file to a Pandas dataframe and remove NaNsdfallstudents = extract_students_by_type(dfallstudents,classyear, studtype) # extract relevant students on programmedfallstudents = remove_unused_columns(dfallstudents)  # remove unused columns (e.g. Admit Term, PSI etc.)df = remove_unnamed_columns(dfallstudents)  # Remove Unnamed columns - this will be used for the outputsids = get_sids(df)          # get student ids for all relevant studentscolumn_names=df.columns.to_list()   # get column namesnstudents = len(sids)                    # Number of students to deal withcounter = 0  # counter for student numberdidnotprocess = ['']  # array for storing which students were actually ignored# loop over SIDsfor anid in sids:    # info    counter += 1    print('Processing student ID {0:s} ({1:d}/{2:d})'.format(anid,counter,nstudents))    dfstudent = get_student(dfallstudents,anid)     # get df for this student including unnamed columns    dfstudent = resequence_units(dfstudent,ignore_courses)  # remove unused units/columns and resequence (codes go in 3rd row)    dfstudent = get_credits(dfstudent,credweights)  # get credits for progression and for weights for calculating averages