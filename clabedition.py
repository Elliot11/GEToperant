from tkinter.filedialog import *
import xlsxwriter
import re
import itertools
#outputfile = asksaveasfilename(title = 'Save output file as', defaultextension='.xlsx', filetypes=(('Excel', '*.xlsx'),('All Files', '*.*')))

def GETpav02(exportfilename, exportstartdate, exportenddate, exportsubject, exportexperiment, exportgroup, exportbox, exportstarttime, exportendtime, exportmsn):
    MPCdatafiles = askopenfilenames(title = 'Select files to import')
    outputdir = askdirectory(title = 'Select directory to save exported files to')
    for f in range(len(MPCdatafiles)):
        Startdate = list()
        Enddate = list()
        Subject = list()
        Experiment = list()
        Group = list()
        Box = list()
        Starttime = list()
        Endtime = list()
        MSN = list()
        A = list()
        B = list()
        C = list()
        D = list()
        E = list()
        F = list()
        G = list()
        H = list()
        I = list()
        J = list()
        K = list()
        L = list()
        M = list()
        N = list()
        O = list()
        P = list()
        Q = list()
        R = list()
        S = list()
        T = list()
        U = list()
        V = list()
        W = list()
        X = list()
        Y = list()
        Z = list()
        Comments = list()

        ### Values will hold the numbers for each array so they can be collected and flattened
        values = list()
        currentarray = ''

        MPC_file = list()
        MPC_file.append(open(MPCdatafiles[f], 'r').readlines())

        ### Begin the for loop that will loop over the data and collect everything into MPC_file
        for i in MPC_file:
            for line in i:
                # Begin by collecting the headers
                # Collect the file names
                if 'File' in line:
                    Filename = line[6:len(line)-1]
                # Collect the start and end dates in ISO 8601 format, correcting for a lack of Y2KCOMPLIANT.
                elif 'Start Date' in line:
                    if len(line) < 22:
                        Startdate.append("20"+line[18:len(line)-1]+"-"+line[12:14]+"-"+line[15:17])
                    else:
                        Startdate.append(line[18:len(line)-1]+"-"+line[12:14]+"-"+line[15:17])
                elif 'End Date' in line:
                    if len(line) < 20:
                        Enddate.append("20"+line[16:len(line)-1]+"-"+line[10:12]+"-"+line[13:15])
                    else:
                        Enddate.append(line[16:len(line)-1]+"-"+line[10:12]+"-"+line[13:15])
                # Similarly, collect subject, experiment, group, box, start time, end time and program name
                elif 'Subject' in line:
                    Subject.append(line[9:len(line)-1])
                elif 'Experiment' in line:
                    Experiment.append(line[12:len(line)-1])
                elif 'Group' in line:
                    Group.append(line[7:len(line)-1])
                elif 'Box' in line:
                    Box.append(line[5:len(line)-1])
                elif 'Start Time' in line:
                    if line[12] == ' ':
                        Starttime.append(line[13:len(line)-1])
                    else:
                        Starttime.append(line[12:len(line)-1])
                elif 'End Time' in line:
                    if line[10] == ' ':
                        Endtime.append(line[11:len(line)-1])
                    else:
                        Endtime.append(line[10:len(line)-1])
                elif 'MSN' in line:
                    MSN.append(line[5:len(line)-1])
                # Check for an array header, if it is present, check if values have been entered into
                # a previous data array. If there are previous data values, flatten the data array and dump them.
                elif len(line) > 1:
                    part_check = re.search(r'\D:', line)
                    if part_check != None:
                        if len(values) > 0:
                            values = list(itertools.chain.from_iterable(values))
                            eval(currentarray).append(values)
                            values = list()
                        ### here we should check for whether the letter has been printed as just a variable.
                        part_checkb = re.search(r'\d', line)
                        if part_checkb != None:
                            currentarray = line[0]
                            values.append(line.split()[1])
                        ### then we should set the beginning of a new array.
                        else:
                            currentarray = line[0]
                    ### this part should then collect data into a new array
                    else:
                        values.append(line.split()[1:])
                    if re.search(r'[\\]', line) != None:
                        Comments.append(line[1:len(line)-1])
                elif len(line) < 1 and len(values) > 0:
                    values = list(itertools.chain.from_iterable(values))
                    eval(currentarray).append(values)
                    values = list()
                elif len(Startdate) > len(Comments):
                        Comments.append(None)

        ### This part will iterate over the CS onset array to get PreCS, CS and PostCS times
        ### Establish totals to export
        PreSessionDelay = list()
        TotalPE = list()
        TotalITI = list()
        TotalPreCS = list()
        TotalCS = list()
        TotalPost = list()
        TotalDuration = list()
        TotalITIDuration = list()
        TotalPreCSDuration = list()
        TotalCSDuration = list()
        TotalPostDuration = list()
        ITIresponsebins = list()
        ITIdurationbins = list()
        PreCSresponsebins = list()
        PreCSdurationbins = list()
        PreCSlatencybins = list()
        CSresponsebins = list()
        CSdurationbins = list()
        CSlatencybins = list()
        PostCSresponsebins = list()
        PostCSdurationbins = list()
        PostCSlatencybins = list()

        ### Establish binned data to export
        ITI_res = list()
        ITI_duration = list()
        PreCS_res = list()
        PreCS_duration = list()
        PreCS_latency = list()
        CS_res = list()
        CS_duration = list()
        CS_latency = list()
        Post_res = list()
        Post_duration = list()
        Post_latency = list()

        for i in range(len(Subject)):
            # Control issues
            CSseconds = float(A[i][2])
            PreSessDelay = float(A[i][0])
            PreCSstarts = list()
            CSstarts = list()
            PostCSstarts = list()

            # Data to extract
            delayresponse = 0
            ITIresponses = list([0])
            ITIdurations = list([0])
            PreCSresponses = list()
            PreCSduration = list()
            PreCSlatency = list()
            CSresponses = list()
            CSduration = list()
            CSlatency = list()
            PostCSresponses = list()
            PostCSduration = list()
            PostCSlatency = list()

            for x in range(1,len(I[i])):
                PreCSstarts.append(float(I[i][x]) - CSseconds)
                CSstarts.append(float(I[i][x]))
                PostCSstarts.append(float(I[i][x]) + CSseconds)

            for k in range(len(CSstarts)):
                PreCSlatency.append(CSseconds)
                CSlatency.append(CSseconds)
                PostCSlatency.append(CSseconds)

                ITIresponses.append(0)
                ITIdurations.append(0)
                PreCSresponses.append(0)
                PreCSduration.append(0)
                CSresponses.append(0)
                CSduration.append(0)
                PostCSresponses.append(0)
                PostCSduration.append(0)
                
            trial = 0
            for response in range(1,len(K[i])):
                ### Check for a response in the delay period
                if float(K[i][response]) < PreSessDelay:
                    delayresponse = delayresponse + 1

                ### Check for a response after the current trial.
                elif float(K[i][response]) >= (CSstarts[-1] + 2 * CSseconds) or float(K[i][response]) >= (CSstarts[trial] + 2 * CSseconds):
                    if trial == len(I[i]) - 2:
                        trial = trial + 1
                        ITIresponses[trial] = ITIresponses[trial] + 1
                        ITIdurations[trial] = ITIdurations[trial] + float(L[i][response])
                    elif trial > len(I[i]) - 2:
                        ITIresponses[trial] = ITIresponses[trial] + 1
                        ITIdurations[trial] = ITIdurations[trial] + float(L[i][response])
                    elif trial < len(I[i]) - 2:
                        while float(K[i][response]) >= (CSstarts[trial] + 2 * CSseconds) and trial < len(I[i]) - 2:
                            trial = trial + 1                          
                        ### Check if this response also needs to be binned
                        if float(K[i][response]) < PreCSstarts[trial]:
                            ITIresponses[trial] = ITIresponses[trial] + 1
                            ITIdurations[trial] = ITIdurations[trial] + float(L[i][response])
                        elif float(K[i][response]) >= PreCSstarts[trial] and float(K[i][response]) < CSstarts[trial]:
                            PreCSresponses[trial] = PreCSresponses[trial] + 1
                            PreCSduration[trial] = PreCSduration[trial] + float(L[i][response])
                            if PreCSlatency[trial] >= CSseconds:
                                PreCSlatency[trial] = float(K[i][response]) - PreCSstarts[trial]
                        elif float(K[i][response]) >= CSstarts[trial] and float(K[i][response]) < PostCSstarts[trial]:
                            CSresponses[trial] = CSresponses[trial] + 1
                            CSduration[trial] = CSduration[trial] + float(L[i][response])
                            if CSlatency[trial] >= CSseconds:
                                CSlatency[trial] = float(K[i][response]) - CSstarts[trial]
                        elif float(K[i][response]) >= PostCSstarts[trial] and float(K[i][response]) < (PostCSstarts[trial] + CSseconds):
                            PostCSresponses[trial] = PostCSresponses[trial] + 1
                            PostCSduration[trial] = PostCSduration[trial] + float(L[i][response])
                            if PostCSlatency[trial] >= CSseconds:
                                PostCSlatency[trial] = float(K[i][response]) - PostCSstarts[trial]
                        elif float(K[i][response]) >= (CSstarts[-1] + 2 * CSseconds):
                            trial = len(I[i]) - 1
                            ITIresponses[trial] = ITIresponses[trial] + 1
                            ITIdurations[trial] = ITIdurations[trial] + float(L[i][response])
                ### Bin the responses normally
                elif float(K[i][response]) < PreCSstarts[trial]:
                    ITIresponses[trial] = ITIresponses[trial] + 1
                    ITIdurations[trial] = ITIdurations[trial] + float(L[i][response])
                elif float(K[i][response]) >= PreCSstarts[trial] and float(K[i][response]) < CSstarts[trial]:
                    PreCSresponses[trial] = PreCSresponses[trial] + 1
                    PreCSduration[trial] = PreCSduration[trial] + float(L[i][response])
                    if PreCSlatency[trial] >= CSseconds:
                        PreCSlatency[trial] = float(K[i][response]) - PreCSstarts[trial]
                elif float(K[i][response]) >= CSstarts[trial] and float(K[i][response]) < PostCSstarts[trial]:
                    CSresponses[trial] = CSresponses[trial] + 1
                    CSduration[trial] = CSduration[trial] + float(L[i][response])
                    if CSlatency[trial] >= CSseconds:
                        CSlatency[trial] = float(K[i][response]) - CSstarts[trial]
                elif float(K[i][response]) >= PostCSstarts[trial] and float(K[i][response]) < (PostCSstarts[trial] + CSseconds):
                    PostCSresponses[trial] = PostCSresponses[trial] + 1
                    PostCSduration[trial] = PostCSduration[trial] + float(L[i][response])
                    if PostCSlatency[trial] >= CSseconds:
                        PostCSlatency[trial] = float(K[i][response]) - PostCSstarts[trial]

            ### Append this data to the lists
            PreSessionDelay.append(list([delayresponse]))
            TotalPE.append(list([sum(ITIresponses) + sum(PreCSresponses) + sum(CSresponses) + sum(PostCSresponses)]))
            TotalITI.append(list([sum(ITIresponses)]))
            TotalPreCS.append(list([sum(PreCSresponses)]))
            TotalCS.append(list([sum(CSresponses)]))
            TotalPost.append(list([sum(PostCSresponses)]))
            TotalDuration.append(list([sum(ITIdurations) + sum(PreCSduration) + sum(CSduration) + sum(PostCSduration)]))
            TotalITIDuration.append(list([sum(ITIdurations)]))
            TotalPreCSDuration.append(list([sum(PreCSduration)]))
            TotalCSDuration.append(list([sum(CSduration)]))
            TotalPostDuration.append(list([sum(PostCSduration)]))
            ITIresponsebins.append(ITIresponses)
            ITIdurationbins.append(ITIdurations)
            PreCSresponsebins.append(PreCSresponses)
            PreCSdurationbins.append(PreCSduration)
            PreCSlatencybins.append(PreCSlatency)
            CSresponsebins.append(CSresponses)
            CSdurationbins.append(CSduration)
            CSlatencybins.append(CSlatency)
            PostCSresponsebins.append(PostCSresponses)
            PostCSdurationbins.append(PostCSduration)
            PostCSlatencybins.append(PostCSlatency)

        fullpath =  outputdir + '/' + Filename.split('\\')[-1] + '.xlsx'

        ### Set up the process to export the data

        Label = list(['Reported TotalPE',
            'Reported ITI PE',
            'Reported PreCS PE',
            'Reported CS PE',
            'Reported PostCS PE',
            'Calculated Total PE',
            'Calculated ITI PE',
            'Calculated PreCS PE',
            'Calculated CS PE',
            'Calculated PostCS PE',
            'Total PE Duration',
            'Total ITI PE Duration',
            'Total PreCS PE Duration',
            'Total CS PE Duration',
            'Total PostCS PE Duration',
            'PE in ITI',
            'PE Duration in ITI',
            'PE in PreCS',
            'PE Duration in PreCS',
            'Latency in PreCS',
            'PE in CS',
            'PE Duration in CS',
            'Latency in CS',
            'PE in PostCS',
            'PE Duration in PostCS',
            'Latency in PostCS',
            'Comment'])
        LabelStartValue= list([None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            None])
        LabelIncrement = list([None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            None])
        ArrayVar = list(['B',
            'B',
            'B',
            'B',
            'B',
            'TotalPE',
            'TotalITI',
            'TotalPreCS',
            'TotalCS',
            'TotalPost',
            'TotalDuration',
            'TotalITIDuration',
            'TotalPreCSDuration',
            'TotalCSDuration',
            'TotalPostDuration',
            'ITIresponsebins',
            'ITIdurationbins',
            'PreCSresponsebins',
            'PreCSdurationbins',
            'PreCSlatencybins',
            'CSresponsebins',
            'CSdurationbins',
            'CSlatencybins',
            'PostCSresponsebins',
            'PostCSdurationbins',
            'PostCSlatencybins',
            'Comments'])
        StartElement = list([16,
            20,
            21,
            22,
            23,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0])
        ArrayIncrement = list([0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            1,
            0])
        StopElement = list([None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None])

        ###

        output = xlsxwriter.Workbook(fullpath)
        output.set_properties({
                'title': 'Batch-Extracted Pav02 Data',
                'subject': 'Animal behaviour',
                'category': 'Raw data',
                'comments': 'Extracted using GEToperant, a Python program using xlrd and xlsxwriter. https://www.github.com/SKhoo'
                })

        mainsheet = output.add_worksheet('GEToperant output')
        mainsheet.set_column('A:A', 23)

        lastrow = -1

        if exportfilename == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Filename')
            mainsheet.write(lastrow, 1, Filename)

        if exportstartdate == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Start Date')
            for i in range(len(Startdate)):
                mainsheet.write(lastrow, i+1, Startdate[i])

        if exportenddate == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'End Date')
            for i in range(len(Enddate)):
                mainsheet.write(lastrow, i+1, Enddate[i])

        if exportsubject == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Subject')
            for i in range(len(Subject)):
                mainsheet.write(lastrow, i+1, Subject[i])

        if exportexperiment == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Experiment')
            for i in range(len(Subject)):
                mainsheet.write(lastrow, i+1, Experiment[i])

        if exportgroup == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Group')
            for i in range(len(Group)):
                mainsheet.write(lastrow, i+1, Group[i])

        if exportbox == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Box')
            for i in range(len(Box)):
                mainsheet.write(lastrow, i+1, float(Box[i]))

        if exportstarttime == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'Start Time')
            for i in range(len(Starttime)):
                mainsheet.write(lastrow, i+1, Starttime[i])

        if exportendtime == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'End Time')
            for i in range(len(Endtime)):
                mainsheet.write(lastrow, i+1, Endtime[i])

        if exportmsn == 1:
            lastrow = lastrow + 1
            mainsheet.write(lastrow, 0, 'MSN')
            for i in range(len(MSN)):
                mainsheet.write(lastrow, i+1, MSN[i])

        for i in range(len(Label)):
            ### This function will loop over the profile. For each label it will check if it is
            ### 1. A single element extraction
            ### 2. A partial array extraction
            ### 3. A full array extraction
            if ArrayIncrement[i] < 1:
                # Single element extraction takes only the label
                lastrow = lastrow + 1
                mainsheet.write(lastrow, 0, Label[i])
                if 'comment' in ArrayVar[i].lower():
                    for k in range(len(Subject)):
                        if k < len(Comments):
                            mainsheet.write(lastrow, k+1, Comments[k])
                        else:
                            mainsheet.write(lastrow, k+1, None)
                else:
                    for k in range(len(Subject)):
                        mainsheet.write(lastrow, k+1, float(eval(ArrayVar[i])[k][StartElement[i]]))
            elif ArrayIncrement[i] > 0:
                if StopElement[i] == None or isinstance(StopElement[i], str):
                    steps = range(StartElement[i], len(max(eval(ArrayVar[i]), key = len)), ArrayIncrement[i])
                elif StopElement[i] > StartElement[i]:
                    steps = range(StartElement[i], StopElement[i] + 1, ArrayIncrement[i])
                for x in steps:
                    lastrow = lastrow + 1
                    for k in range(len(Subject)):
                        if LabelIncrement[i] != None and LabelIncrement[i] > 0:
                            mainsheet.write(lastrow, 0, Label[i] + ' ' + str(LabelStartValue[i] + x * LabelIncrement[i]))
                        else:
                            mainsheet.write(lastrow, 0, Label[i])
                        if x < len(eval(ArrayVar[i])[k]):
                            mainsheet.write(lastrow, k+1, float(eval(ArrayVar[i])[k][x]))
                        else:
                            mainsheet.write(lastrow, k+1, None)

        output.close()

