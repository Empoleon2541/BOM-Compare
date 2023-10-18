import pandas as pd
import os
import glob
import openpyxl as xl
import datetime
import tkinter as tk
import shutil
import numpy as np

def disable_event():
    pass
def close_window():
    window.destroy()
def click(event):
    UIEntry.delete(0,100)
    UIEntry.update

###################################### Functions for cross Preliminary work ##########################################################################################################################################################
def getPath():
    # function to be called when the 'Analyze' button is clicked
    # main algorithm for the preliminary work of the program

    global prereqCheckDesignator, path, mainFolderFileList, referenceTable, BomTable
    statusOut['text'] = "Analyzing and matching BOMs provided"
    statusOut.update()
    path = UIEntry.get()

    if not os.path.exists(path):
        statusOut['text'] = 'Please provide a valid path'
        return

    newFolderPath = path + r'\New'
    oldFolderPath = path + r'\Old'
    
    if not os.path.exists(newFolderPath) or not os.path.exists(oldFolderPath):
        statusOut['text'] = 'Please follow the required folder and file format then try again'
        return
    
    mainFolderFileList = glob.glob(os.path.join(path, '*.xlsx'))
    newFolderFiles = glob.glob(os.path.join(newFolderPath, '*.xlsx'))
    oldFolderFiles = glob.glob(os.path.join(oldFolderPath, '*.xlsx'))

    if len(newFolderFiles) == 0:
        statusOut['text'] = 'New BOM folder is empty'
        return
        
    if len(oldFolderFiles) == 0:
        statusOut['text'] = 'Old BOM folder is empty'
        quit()
    
    try:
        
        newList = summary(newFolderFiles)
        oldList = summary(oldFolderFiles)

    except PermissionError as e:
    
        statusOut['text'] = 'Please Close all excel files within the directory then try again'
        quit()
        
    referenceTable, BomTable = generateComparisonTable(newList,oldList)

    summaryPath = path + r'\Summary.xlsx'

    wb = xl.Workbook()
    wb.save(summaryPath)

    with pd.ExcelWriter(summaryPath,mode="a", engine="openpyxl", if_sheet_exists="overlay",) as writer:
            
            referenceTable.to_excel(writer, sheet_name = 'Comparison Table')
            BomTable.to_excel(writer, sheet_name = 'BOM list')
                    
            ws = writer.sheets['Comparison Table']
            wb = writer.sheets['BOM list']

            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 30

            wb.column_dimensions['A'].width = 20
            wb.column_dimensions['B'].width = 30
            wb.column_dimensions['C'].width = 180
            wb.column_dimensions['D'].width = 50

            wx = writer.book
            std = writer.sheets['Sheet']
            wx.remove(std) 

    prereqCheckDesignator = prereqCheck(referenceTable, BomTable)

    referenceTable =  referenceTable.reset_index()
    BomTable =  BomTable.reset_index()
def summary(filelist):
    #summarizes the available files for each folder
    fileListDataFrame = pd.DataFrame(columns = ('Item_ID','Version','Device','FilePath'))

    for excelfile in filelist:

        loadedExcel = pd.ExcelFile(excelfile)
        bom = pd.read_excel(loadedExcel, sheet_name = 0)
        deviceName = (bom['Unnamed: 1'][4])
        
        deviceEntry = pd.DataFrame([deviceName.split('/')[0],deviceName.split('/')[1][:2],deviceName.split('- ')[1][:-1],bom.columns[0],excelfile]).T
        deviceEntry = deviceEntry.set_axis(['Item_ID','Version','Device','FileType','FilePath'], axis = 'columns')
        fileListDataFrame = pd.concat([fileListDataFrame,deviceEntry])
        
    return fileListDataFrame
def generateComparisonTable(newList,oldList):
    # generates the comparison table by matching old boms to new boms
    # would first try to look for same assembly PN i.e. 121212-01/02 ----> 121212-01/01
    # if version 01, would look for closest name match by iteration
    # will look for closest match by iterating through the available assembly names and deleting the last word from the name of the new bom
    # i.e. will look for name of: X308 NiPm Alt SMPS -------> X308 NiPm Alt --------> X308 Nipm; until a match is found
    completeList = pd.concat([newList,oldList]).reset_index(drop=True)
    comparisonTable = newList.drop(['FileType','FilePath'],axis='columns').reset_index(drop=True)
    oldList = oldList.reset_index(drop=True)
    comparisonTable['Reference_ID'] = 'Reference Not Found'
    comparisonTable['Reference_Device'] = ''

    for i in comparisonTable.index:
        
        Item_ID = comparisonTable['Item_ID'][i]
        comparisonTable['Item_ID'][i] = Item_ID + '/' + comparisonTable['Version'][i]
        if Item_ID in oldList['Item_ID'].values:

            comparisonTable['Reference_ID'][i] = (Item_ID + '/'+ oldList['Version'][oldList[oldList['Item_ID']==Item_ID].index.values[0]])
            comparisonTable['Reference_Device'][i] = oldList['Device'][oldList[oldList['Item_ID']==Item_ID].index.values[0]]

        elif comparisonTable['Version'][i] == '01':

            desc = (' ').join(comparisonTable['Device'][i].split(' ')[:-1])
            
            while desc not in oldList['Device'].values and desc != '':
                desc = (' ').join(desc.split(' ')[:-1])

            if desc != '':
                comparisonTable['Reference_ID'][i] = (oldList['Item_ID'][oldList[oldList['Device']==desc].index.values[0]] + '/'+ oldList['Version'][oldList[oldList['Device']==desc].index.values[0]])
                comparisonTable['Reference_Device'][i] = desc
    
    for i in completeList.index:
        completeList['Item_ID'][i] = completeList['Item_ID'][i] + '/' +completeList['Version'][i]

    completeList = completeList.drop(completeList.columns[1],axis = 1).set_index('Item_ID')
    comparisonTable = comparisonTable.drop(comparisonTable.columns[1],axis = 1).set_index('Item_ID')

    return comparisonTable, completeList
def prereqCheck(referenceList, BomList):
    #checks if all files are correct and all New Boms have a match
    prereqCheckDesignator = 0
    if 'Reference Not Found' in referenceList['Reference_ID'].values:
        statusOut['text'] = 'Please fix "Reference Not Found" issues first'
        prereqCheckDesignator==1

    if not all(fileType == 'BOM Report - Engineering       ' for fileType in BomList['FileType'].values):
        statusOut['text'] = 'Please ensure all files are generated as "BOM Report - Engineering"'
        prereqCheckDesignator==1

    if prereqCheckDesignator == 0:
        
        statusOut['text'] = 'Please check Summary file for verification'
    return prereqCheckDesignator 

######################################## Functions for BOM Comparison ##########################################################################################################################################################
def fetchData(referenceList, bomList, i):
    #fetches the data by reading the matching BOMs from the referenceList and looking up for the filepath for each bom from the bomList
    Item_ID = referenceList['Item_ID'][i]
    Ref_ID = referenceList['Reference_ID'][i]
    
    new_path = bomList['FilePath'][bomList[bomList['Item_ID']==Item_ID].index.values[0]]
    ref_path = bomList['FilePath'][bomList[bomList['Item_ID']==Ref_ID].index.values[0]]

    newBom = pd.read_excel(new_path, sheet_name = 0)
    oldBom = pd.read_excel(ref_path, sheet_name = 0)

    return newBom, oldBom, Item_ID[:-3]
def cleaner(db):
    # used to clean BOMs 
    db = db.iloc[16:]
    db = db.drop(db.columns[[0,5,6,8,9,11,12]],axis = 1)
    db = db.set_axis(['Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts'], axis = 'columns')
    db = db.fillna('')
    db = db.groupby(['Item_ID','Revision','Full_Name','Description','Alternate_Parts'])['Designator'].apply(','.join).reset_index()
    return db
def alternatePartCheck(newBom,ind,addedAlternatePartsOut,removedAlternatePartsOut,newAlternatePartsList,oldAlternatePartsList):
    # creates two different dataframes that includes the added, and removed, alternate parts
    newAlternatePartStorage = newAlternatePartsList.split('\n')
    oldAlternatePartStorage = oldAlternatePartsList.split('\n')

    Alt_New = [x for x in newAlternatePartStorage if not x in oldAlternatePartStorage ]
    Alt_Rem = [x for x in oldAlternatePartStorage if not x in newAlternatePartStorage ]
                                            
    for j in Alt_New: 

        if j != '':       
            addedAlternatePartsOut = pd.concat([addedAlternatePartsOut, newBom.iloc[ind].to_frame().T])
            addedAlternatePartsOut = addedAlternatePartsOut.reset_index(drop=True) 
            addedAlternatePartsOut.at[addedAlternatePartsOut.last_valid_index(),'Alternate_Parts'] = j

    for j in Alt_Rem:
                    
        if j != '':    
            removedAlternatePartsOut = pd.concat([removedAlternatePartsOut, newBom.iloc[ind].to_frame().T])
            removedAlternatePartsOut = removedAlternatePartsOut.reset_index(drop=True) 
            removedAlternatePartsOut.at[removedAlternatePartsOut.last_valid_index(),'Alternate_Parts'] = j
    
    return addedAlternatePartsOut, removedAlternatePartsOut
def designatorCheck(newBom,ind,Item_ID,removedDesignatorOut,addedDesignatorOut,newDesignatorList,oldDesignatorList):
    # creates two different dataframes that includes the added and removed part designations
    newDesignatorStorage = ungroup(newDesignatorList)
    oldDesignatorStorage = ungroup(oldDesignatorList)

    Alternates = newBom['Alternate_Parts'][newBom[newBom['Item_ID']==Item_ID].index.values[0]].split('\n')

    d_add = ','.join([x for x in newDesignatorStorage if not x in oldDesignatorStorage])
    d_rem = ','.join([x for x in oldDesignatorStorage if not x in newDesignatorStorage])

    if d_add != '':
        for j in Alternates:
            addedDesignatorOut = pd.concat([addedDesignatorOut, newBom.iloc[ind].to_frame().T])
            addedDesignatorOut.at[addedDesignatorOut.last_valid_index(),'Designator'] = d_add
            addedDesignatorOut = addedDesignatorOut.reset_index(drop=True) 
            addedDesignatorOut.at[addedDesignatorOut.last_valid_index(),'Alternate_Parts'] = j
    
    if d_rem != '':
        for j in Alternates:
            removedDesignatorOut = pd.concat([removedDesignatorOut, newBom.iloc[ind].to_frame().T])
            removedDesignatorOut.at[removedDesignatorOut.last_valid_index(),'Designator'] = d_rem
            removedDesignatorOut = removedDesignatorOut.reset_index(drop=True) 
            removedDesignatorOut.at[removedDesignatorOut.last_valid_index(),'Alternate_Parts'] = j
    
    return removedDesignatorOut, addedDesignatorOut
def ungroup(desig):
    # ungroups dashes 
    # mainly used for designations i.e. R1-R3 ---> R1,R2,R3
    desig_var = desig.split(',')
    ungrouped = []

    for i in desig_var:
        
        if '-' in i:
            designations = i.split('-')
            comp = designations[1][0]

            for j in range (0,2):
                designations[j]= int(designations[j][1:])

            for j in range(designations[0],designations[1]+1):
                ungrouped.append(str(comp+str(j)))

        else:
            ungrouped.append(i)
            
    return(ungrouped) 

############################## Functions for cross referencing with the Reference File ##########################################################################################################################################################
def addAlternatePartDescription(alternatePartList,partsListReference):
    # adds part description for alternate parts (would only work if reference file is provided)
    for i in alternatePartList.index:
        
        PN = alternatePartList["Alternate_Parts"][i]
        if PN in partsListReference['Part_Number'].values:
            desc = (partsListReference['Manufacturer'][partsListReference[partsListReference['Part_Number']==PN].index.values[0]] + ' ' + partsListReference['Manufacturer_PN'][partsListReference[partsListReference['Part_Number']==PN].index.values[0]] + partsListReference['Description'][partsListReference[partsListReference['Part_Number']==PN].index.values[0]])
            alternatePartList['Alternate_Part_Full_Description'][i] = desc 
        
    return(alternatePartList)
def extractReference(referencePath):
    # function to extract the data from the reference file
    xl = pd.ExcelFile(referencePath)
    changeList = pd.read_excel(xl, sheet_name = 0)
    changeList = changeList.iloc[7:]
    changeList = changeList.set_axis(['partNumber','addedParts','removedParts','alternatePartMain','alternatePartAdded','alternatePartRemoved','partDesignationMain','partDesignationAdded','partDesignationRemoved'],axis = 'columns').reset_index(drop=True)
    changeList['partNumber'] = changeList['partNumber'].fillna(method = 'ffill')

    partList = pd.read_excel(xl, sheet_name = 1 )
    partList = partList[4:]
    partList = partList.set_axis(['Part_Number','Manufacturer','Manufacturer_PN','Description'], axis = 'columns').reset_index(drop=True)
    partList = partList.fillna('')

    return changeList, partList
def extractChanges(changeList, assemblyPN):   
    # extract changes from the changeList for each assembly
    changeList = changeList[changeList['partNumber'] == assemblyPN].reset_index(drop=True)
    newParts = changeList['addedParts'].dropna().values.tolist()
    removedParts = changeList['removedParts'].dropna().values.tolist()
    alternatePartAdded = [changeList['alternatePartMain'][i] + ',' + changeList['alternatePartAdded'][i] for i in changeList.index if (changeList['alternatePartMain'][i]==changeList['alternatePartMain'][i] and changeList['alternatePartAdded'][i]==changeList['alternatePartAdded'][i])] 
    alternatePartRemoved = [changeList['alternatePartMain'][i] + ',' + changeList['alternatePartRemoved'][i] for i in changeList.index if (changeList['alternatePartMain'][i]==changeList['alternatePartMain'][i] and changeList['alternatePartRemoved'][i]==changeList['alternatePartRemoved'][i])] 
    partDesignationAdded = [changeList['partDesignationMain'][i] + ',' + changeList['partDesignationAdded'][i] for i in changeList.index if (changeList['partDesignationMain'][i]==changeList['partDesignationMain'][i] and changeList['partDesignationAdded'][i]==changeList['partDesignationAdded'][i])] 
    partDesignationRemoved = [changeList['partDesignationMain'][i] + ',' + changeList['partDesignationRemoved'][i] for i in changeList.index if (changeList['partDesignationMain'][i]==changeList['partDesignationMain'][i] and changeList['partDesignationRemoved'][i]==changeList['partDesignationRemoved'][i])] 
    
    return newParts, removedParts, alternatePartAdded, alternatePartRemoved, partDesignationAdded, partDesignationRemoved
def referenceCheck(changeStatus,IssueCounter,changeArray,referenceArray,changeType):
    # checks each change and cross reference it from the listed changes
    if changeType == 'alternatePart':
        changesToCheck = [changeArray['Item_ID'][i] + ',' + changeArray['Alternate_Parts'][i] for i in changeArray.index] 
    elif changeType == 'designation':
        changesToCheck = [changeArray['Item_ID'][i] + ',' + changeArray['Designator'][i] for i in changeArray.index] 
    elif changeType == 'partsList':
        changesToCheck = changeArray['Item_ID'].values.tolist()
    for i in changesToCheck:
        if i in referenceArray:
            changeStatus.append('Okay')
        else:
            changeStatus.append('Not in reference file')
            IssueCounter += 1
    for i in referenceArray:
        if i not in changesToCheck:
            changeStatus.append('Missing from detected changes')
            IssueCounter += 1
            if changeType == 'partsList':
                changeArray = pd.concat([changeArray, pd.DataFrame([i,'','','','','','',np.nan],index=['Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts','Alternate_Part_Full_Description','Remarks']).T]) 
            elif changeType == 'alternatePart':
                changeArray = pd.concat([changeArray, pd.DataFrame([i.split(',')[0],'','','','',i.split(',')[1],'',np.nan],index=['Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts','Alternate_Part_Full_Description','Remarks']).T])
            elif changeType == 'designation':
                changeArray = pd.concat([changeArray, pd.DataFrame([i.split(',')[0],'','','',i.split(',')[1],'','',np.nan],index=['Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts','Alternate_Part_Full_Description','Remarks']).T]) 
    return(changeStatus,IssueCounter,changeArray)            

############################################## Main Program Algorithm ##########################################################################################################################################################
def compare():
    # Main Program. Would be called by 'Compare' Button on UI
    if (prereqCheckDesignator == 1):
        statusOut['text'] = 'Please run/rerun analysis first'
        return 
    window.protocol("WM_DELETE_WINDOW", disable_event)
    now = str(datetime.datetime.now())[:-7]
    now = now.replace(':','.')
    filepath = path + '\\' + now + '.xlsx'
    referencePath = path + r'\Reference.xlsx'
    referenceDesignator = 0
    cnt = 1 #counter used to copy the summary file to the output file//so it would only copy once and with atleast one BOM comparison successfully compared
    totalDoneCounter = 0
    summaryPath = path + r'\Summary.xlsx'

    # checks if reference file exists
    if referencePath in mainFolderFileList:
        
        try:

            changeList, partList = extractReference(referencePath)
            assemblyIssueCounter = []

        except PermissionError as e:
        
            statusOut['text'] = 'Please close the reference excel file and try again'
            
       
        referenceDesignator = 1
    
    emptyRow = pd.Series(['','','','','','',''],index=['Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts','Alternate_Part_Full_Description'])
    
    for index in referenceTable.index:


        addedAlternateParts = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts','Alternate_Part_Full_Description'))
        removedAlternateParts = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts','Alternate_Part_Full_Description'))
        addedParts = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts'))
        removedParts = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts'))
        addedDesignator = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts'))
        removedDesignator = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts'))
        updates = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts'))
        final = pd.DataFrame(columns = ('Item_ID','Revision','Full_Name','Description','Designator','Alternate_Parts', 'Remarks','Alternate_Part_Full_Description'))
        
        newBOM, oldBOM, assemblyName = fetchData(referenceTable, BomTable, index)
        statusOut['text'] = (assemblyName + ' BOM Comparison: running....')
        statusOut.update()
        newBOM = cleaner(newBOM)
        oldBOM = cleaner(oldBOM)
        
        for ind in newBOM.index:
            
            Item_ID = newBOM['Item_ID'][ind]
            
            if Item_ID in oldBOM['Item_ID'].values:
                
                alternatePartsNew = newBOM['Alternate_Parts'][newBOM[newBOM['Item_ID']==Item_ID].index.values[0]]
                alternatePartsOld = oldBOM['Alternate_Parts'][oldBOM[oldBOM['Item_ID']==Item_ID].index.values[0]]

                designatorNew = newBOM['Designator'][newBOM[newBOM['Item_ID']==Item_ID].index.values[0]]
                designatorOld = oldBOM['Designator'][oldBOM[oldBOM['Item_ID']==Item_ID].index.values[0]]
                
                RevisionNew = newBOM['Revision'][newBOM[newBOM['Item_ID']==Item_ID].index.values[0]]
                RevisionOld = oldBOM['Revision'][oldBOM[oldBOM['Item_ID']==Item_ID].index.values[0]]

                if alternatePartsNew != alternatePartsOld:

                    addedAlternateParts,removedAlternateParts = alternatePartCheck(newBOM, ind, addedAlternateParts,removedAlternateParts,alternatePartsNew,alternatePartsOld)
                    
                if designatorNew != designatorOld:
                                                                                
                    removedDesignator, addedDesignator = designatorCheck(newBOM, ind, Item_ID, removedDesignator, addedDesignator,designatorNew,designatorOld)

                if RevisionNew != RevisionOld:
                    
                    updates = pd.concat([updates, newBOM.iloc[ind].to_frame().T])

                oldBOM = oldBOM.drop(oldBOM[oldBOM['Item_ID']==Item_ID].index.values[0])
            

            else:
                
                #sorts and tabulates the new parts with individual tabulation per alternate part
                altParts = newBOM['Alternate_Parts'][newBOM[newBOM['Item_ID']==Item_ID].index.values[0]].split('\n')
                for i in altParts:
                    addedParts = pd.concat([addedParts, newBOM.iloc[ind].to_frame().T])
                    addedParts = addedParts.reset_index(drop=True) 
                    addedParts.at[addedParts.last_valid_index(),'Alternate_Parts'] = i
            
        if referenceDesignator == 1:
                    
            addedAlternateParts = addAlternatePartDescription(addedAlternateParts, partList)
            removedAlternateParts = addAlternatePartDescription(removedAlternateParts, partList)
            
        if not oldBOM.empty:

            #sorts and tabulates the new parts with individual tabulation per alternate part
            oldBOM = oldBOM.reset_index(drop=True) 

            for ind, row in oldBOM.iterrows():
                alt_parts = row['Alternate_Parts'].split('\n')
                for j in alt_parts:
                    removedParts = pd.concat([removedParts, oldBOM.iloc[ind].to_frame().T])
                    removedParts = removedParts.reset_index(drop=True) 
                    removedParts.at[removedParts.last_valid_index(),'Alternate_Parts'] = j

        if referenceDesignator == 1:
            newPartsRef, removedPartsRef, alternatePartAddedRef, alternatePartRemovedRef, partDesignationAddedRef, partDesignationRemovedRef = extractChanges(changeList,assemblyName)
            changeStatus = []
            IssueCounter = 0
            # changeStatus is a sorted list for the status of each change. it will then be appended to the end of the final output after each assembly
            # IssueCounter counts the issues per assembly that will be appended to the summary page to notify the user which assemblies have an issue


        final = pd.concat([final, updates])
        final['Remarks'] = final['Remarks'].fillna('Part Updated') 
        for i in updates.index:
            changeStatus.append('')
        if referenceDesignator == 1:
            changeStatus, IssueCounter, addedAlternateParts = referenceCheck(changeStatus,IssueCounter,addedAlternateParts,alternatePartAddedRef,'alternatePart')
        final = pd.concat([final, addedAlternateParts])
        final['Remarks'] = final['Remarks'].fillna('Alternate Part Added')        
        if referenceDesignator == 1:
            changeStatus, IssueCounter, removedAlternateParts= referenceCheck(changeStatus,IssueCounter,removedAlternateParts,alternatePartRemovedRef,'alternatePart')
        final = pd.concat([final, removedAlternateParts])
        final['Remarks'] = final['Remarks'].fillna('Alternate Part Removed')        
        if referenceDesignator == 1:
            changeStatus, IssueCounter, addedDesignator = referenceCheck(changeStatus,IssueCounter,addedDesignator,partDesignationAddedRef,'designation')
        final = pd.concat([final, addedDesignator])
        final['Remarks'] = final['Remarks'].fillna('New Part Designation Added')        
        if referenceDesignator == 1:
            changeStatus, IssueCounter, removedDesignator = referenceCheck(changeStatus,IssueCounter,removedDesignator,partDesignationRemovedRef,'designation')
        final = pd.concat([final, removedDesignator])
        final['Remarks'] = final['Remarks'].fillna('Old Part Designation Removed')        
        if referenceDesignator == 1:
            changeStatus, IssueCounter, addedParts = referenceCheck(changeStatus,IssueCounter,addedParts,newPartsRef,'partsList')
        final = pd.concat([final, addedParts])
        final['Remarks'] = final['Remarks'].fillna('New Part Added')        
        if referenceDesignator == 1:
            changeStatus, IssueCounter, removedParts = referenceCheck(changeStatus,IssueCounter,removedParts,removedPartsRef,'partsList')
        final = pd.concat([final, removedParts])
        final['Remarks'] = final['Remarks'].fillna('Old Part Removed')        
        final = final.set_index('Item_ID')
        if referenceDesignator == 1:
            print(final)
            final.insert(6,'Change Status',changeStatus)
            if IssueCounter==0:
                assemblyIssueCounter.append('Okay: No Error found')
            else:
                assemblyIssueCounter.append(str(IssueCounter)+ ' Errors found')
        if final.empty:
                
            final = pd.concat([final, emptyRow.to_frame().T])
            final['Remarks'][0] = 'No differences were observed'

        if cnt == 1: 
                shutil.copy2(summaryPath, filepath)
                cnt = 0
                    
        with pd.ExcelWriter(filepath,mode="a", engine="openpyxl", if_sheet_exists="overlay",) as writer:
            
            final.to_excel(writer, sheet_name = assemblyName)

            wb = writer.book
            
            ws = writer.sheets[assemblyName]

            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 10
            ws.column_dimensions['C'].width = 50
            ws.column_dimensions['D'].width = 80
            ws.column_dimensions['E'].width = 15
            ws.column_dimensions['F'].width = 20
            ws.column_dimensions['G'].width = 30
            if referenceDesignator == 1:
                ws.column_dimensions['H'].width = 30
                ws.column_dimensions['I'].width = 100
            else:
                ws.column_dimensions['H'].width = 100

            ws.protection.sheet = True
            ws.protection.password = '148331'
        
        statusOut['text'] = (assemblyName + ' BOM Comparison: Done!                   ')
        statusOut.update()
        totalDoneCounter += 1
    if referenceDesignator == 1:
        referenceTable.insert(4,'Status',assemblyIssueCounter)
        with pd.ExcelWriter(filepath,mode="a", engine="openpyxl", if_sheet_exists="overlay",) as writer:
            
            referenceTable.set_index('Item_ID').to_excel(writer, sheet_name = 'Comparison Table')
                    
            ws = writer.sheets['Comparison Table']

            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 30
            ws.column_dimensions['E'].width = 20



    statusOut['text'] = ('Results Generated. ' + str(totalDoneCounter) + ' BOMs compared.')
    window.protocol("WM_DELETE_WINDOW", close_window)



################################ Everything below accounts for the UI of the Program ##########################################################################################################################################################
x = 1
window = tk.Tk()
window.title('BOM Compare')
frame = tk.Frame(
    master=window,
    width = 345,
    height = 80
)
frame.pack()

UIEntry = tk.Entry(master = frame, width=50)
UIEntry.insert(0, "Enter folder path here")

button1 = tk.Button(
    master = frame,
    text = 'Analyze',
    width = 20,
    command = getPath
    )

button2 = tk.Button(
    master = frame,
    text = 'Compare',
    width = 20,
    command = compare)


button1.place(x=20, y=40)
UIEntry.place(x = 20, y = 10)
button2.place(x=175, y=40)

statusFrame = tk.Frame(master = window)
statusLabel = tk.Label(
    master = statusFrame,
    text = 'Status: '
)

statusOut = tk.Label(
    master = statusFrame,
    text = ''
)

statusLabel.grid(row = 0, column = 0)
statusOut.grid(row = 0, column = 1)

statusFrame.pack()
clicked = UIEntry.bind('<Button-1>', click)

window.mainloop()





    