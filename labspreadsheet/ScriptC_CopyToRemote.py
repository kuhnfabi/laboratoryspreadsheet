#xDemo = True
xDemo = False

import numpy as np
import os
import pandas
from pathlib import Path
from shutil import copyfile
from tools import GetMetaData
import datetime

TimeNow = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

if xDemo:
    folderLocal = './demo/'
else:
    folderLocal = './production/'
    
folderMeta = folderLocal+'meta/'
folderOutput = folderLocal+'local_output/'
metafile = 'MetaData.xlsx'

# =======================================
#    1. Collect meta-information
# =======================================

metapath = os.path.join(Path(folderMeta),Path(metafile))
dfSITE, dfVAR, lstSampleMethod, lstSampleFilter = GetMetaData(metapath)

dfProject = pandas.read_excel(io =metapath, engine="openpyxl",sheet_name ='file',header=7,dtype='str',keep_default_na=False).dropna(axis=0,how='all')
dfProject.sort_values(by=['Year', 'Project'],inplace=True)

dfCentral = pandas.read_excel(io =metapath, engine="openpyxl", sheet_name ='file',header=3,nrows=2,dtype='str',keep_default_na=False).dropna(axis=0,how='all')
folderArchive = Path(dfCentral.Folder[1])
folderCentral = Path(dfCentral.Folder[0])

# =======================================
#    2. Copy files
# =======================================

years = sorted(dfProject.Year.unique())
years_available = np.zeros(len(years), dtype=bool)

for i, year in enumerate(years):
    
    if datetime.datetime.now().year<=np.float(year):
        folder_samples = folderCentral
    else:
        folder_samples = folderArchive
        
    filenameSample = 'LabSamples'+str(year)+'.xlsx'
    filepathSampleArchive =  os.path.join(Path(folderArchive),'automatic_backup/rxv'+TimeNow+'_'+filenameSample)
    filepathSampleRemote = os.path.join(folder_samples,filenameSample)
    filepathSampleLocal =  os.path.join(Path(folderOutput),filenameSample)
    
    print('Target file location: '+filepathSampleRemote)
    print('Archive file location: '+filepathSampleArchive)
    if os.path.isfile(filepathSampleRemote):
        
        try:
            copyfile(filepathSampleRemote, filepathSampleArchive)
            print('--- Archival of old file successful')
        
            try:
                copyfile(filepathSampleLocal,filepathSampleRemote)
                print('--- Overwriting of old file with new file successful')
            except:
                print('--- WARNING: Overwriting of old file with new file failed')
        except:
            print('--- WARNING: Archival of old file failed, overwriting of file is skipped')
    else:
        print('--- No old copy of file has been found - will create new one')
        try:
            copyfile(filepathSampleLocal,filepathSampleRemote)
            print('--- Copying new file successful')
        except:
            print('--- WARNING: Copying new file failed')
        
for row in dfProject.itertuples(index=True, name='Pandas'):
    
    strSheetName = str(row.Project)
    filenames = ['LabSamples'+str(row.Year)+'_'+strSheetName+'_extra'+'.xlsx',\
                 'LabResults'+str(row.Year)+'_'+strSheetName+'_extra'+'.xlsx',\
                 'LabResults'+str(row.Year)+'_'+strSheetName+'.xlsx']
    
    for filename in filenames:
        filepathArchive = os.path.join(Path(folderArchive),'automatic_backup/rxv'+TimeNow+'_'+filename)
        filepathRemote = os.path.join(Path(row.Folder),filename)
        filepathLocal = os.path.join(Path(folderOutput),filename)
        print('Target file location: '+filepathRemote)
        print('Archive file location: '+filepathArchive)
        if os.path.isfile(filepathRemote):
            try:
                copyfile(filepathRemote, filepathArchive)
                print('--- Archival of old file successful')
        
                try:
                    copyfile(filepathLocal,filepathRemote)
                    print('--- Overwriting of old file with new file successful')
                except:
                    print('--- WARNING: Overwriting of old file with new file failed')
            except:
                print('--- WARNING: Archival of old file failed, overwriting of file is skipped')
        else:
            print('--- No old copy of file has been found - will create new one')
            try:
                copyfile(filepathLocal,filepathRemote)
                print('--- Copying new file successful')
            except:
                print('--- WARNING: Copying new file failed')
            
            
# =======================================
#    3. Finalization
# =======================================
print('Done!')