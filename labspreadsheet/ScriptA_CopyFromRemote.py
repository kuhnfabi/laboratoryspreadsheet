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
folderInput = folderLocal+'local_input/'
metafile = 'MetaData.xlsx'

# =======================================
#    1. Collect meta-information
# =======================================

metapath = os.path.join(Path(folderMeta),Path(metafile))
print(metapath)
dfSITE, dfVAR, lstSampleMethod, lstSampleFilter = GetMetaData(metapath)

dfProject = pandas.read_excel(io =metapath, engine="openpyxl",sheet_name ='file',header=7,dtype='str',keep_default_na=False,na_values = {""}).dropna(axis=0,how='all')

dfCentral = pandas.read_excel(io =metapath, engine="openpyxl",sheet_name ='file',header=3,nrows=2,dtype='str',keep_default_na=False,na_values = {""}).dropna(axis=0,how='all')
folderArchive = Path(dfCentral.Folder[1])
folderCentral = Path(dfCentral.Folder[0])

# =======================================
#    2. Delete existing files
# =======================================

[ os.remove(os.path.join(Path(folderInput),f)) for f in os.listdir(Path(folderInput)) if f.endswith(".xlsx") ];

# =======================================
#    3. Copy files
# =======================================

years = sorted(dfProject.Year.unique())
years_available = np.zeros(len(years), dtype=bool)

for i, year in enumerate(years):
    
    filenameSample = 'LabSamples'+str(year)+'.xlsx'
    filepathSampleArchive =  os.path.join(Path(folderArchive),'automatic_backup/rxv'+TimeNow+'_'+filenameSample)
    filepathSampleLocal =  os.path.join(Path(folderInput),filenameSample)
    filepathSampleRemote = os.path.join(Path(folderCentral),filenameSample)
    print(filepathSampleArchive)
    
    try:
        copyfile(filepathSampleRemote, filepathSampleArchive)
        print('Archived: '+filepathSampleRemote)
        years_available[i] = True
    except:
        print('WARNING - Did not archive: '+filepathSampleRemote)
        
    try:
        copyfile(filepathSampleRemote, filepathSampleLocal)
        print('Copied: '+filepathSampleRemote)
        years_available[i] = True
    except:
        print('WARNING - Did not copy: '+filepathSampleRemote)
        

for i, year in enumerate(years):
    
    filenameSample = 'LabSamples'+str(year)+'.xlsx'
    filepathSampleArchive =  os.path.join(Path(folderArchive),'automatic_backup/rxv'+TimeNow+'_'+filenameSample)
    filepathSampleLocal =  os.path.join(Path(folderInput),filenameSample)
    filepathSampleRemote = os.path.join(Path(folderArchive),filenameSample)
    
    try:
        copyfile(filepathSampleRemote, filepathSampleArchive)
        print('Archived: '+filepathSampleRemote)
        years_available[i] = True
    except:
        print('WARNING - Did not archive: '+filepathSampleRemote)
        
    try:
        copyfile(filepathSampleRemote, filepathSampleLocal)
        print('Copied: '+filepathSampleRemote)
        years_available[i] = True
    except:
        print('WARNING - Did not copy: '+filepathSampleRemote)
    
for row in dfProject.itertuples(index=True, name='Pandas'):
    
    filenames = ['LabSamples'+str(row.Year)+'_'+str(row.Project)+'_extra'+'.xlsx',\
                 'LabResults'+str(row.Year)+'_'+str(row.Project)+'_extra'+'.xlsx',\
                 'LabResults'+str(row.Year)+'_'+str(row.Project)+'.xlsx']
    
    for filename in filenames:
        filepathArchive = os.path.join(Path(folderArchive),'automatic_backup/rxv'+TimeNow+'_'+filename)
        filepathRemote = os.path.join(Path(row.Folder),filename)
        filepathLocal = os.path.join(Path(folderInput),filename)
        try:
            copyfile(filepathRemote, filepathArchive)
            print('Archived: '+filepathRemote)
            years_available[i] = True
        except:
            print('WARNING - Did not archive: '+filepathRemote)
            
        try:
            copyfile(filepathRemote, filepathLocal)
            print('Obtained: '+filepathRemote)
            years_available[i] = True
        except:
            print('WARNING - Did not copy: '+filepathRemote)


# =======================================
#    4. Finalization
# =======================================
print('Done!')