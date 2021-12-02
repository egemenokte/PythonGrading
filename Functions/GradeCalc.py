import numpy as np
import pandas as pd
import io
def CalculateGrade(uploaded):
    filename = next(iter(uploaded))
    xlsx_file = io.BytesIO(uploaded.get(filename))
    Students = pd.read_excel(xlsx_file, 'Students')
    Percentages = pd.read_excel(xlsx_file, 'Percentages')
    Letters = pd.read_excel(xlsx_file, 'Letters')
    SumDict={}
    AssignmentList=list(Students.columns)[4:]
    for i in range(len(AssignmentList)):
        Type = AssignmentList[i].split()[0]
        if Type in SumDict:
            if len(AssignmentList[i].split())>1:
                Number=int(AssignmentList[i].split()[1])
                if Number>SumDict[Type]:
                    SumDict[Type]=Number
        else:
            SumDict[Type]=1
    TypeDict={}

    for i in range(len(Percentages)-1):
        TypeDict[Percentages.loc[i,'Type']]={}
        for j in range(1,len(Percentages.columns)):
            Type=Percentages.columns[j]
            TypeDict[Percentages.loc[i,'Type']][Type]=[]
            if isinstance(Percentages.iloc[i,j], str):
                lst_int = [int(x) for x in Percentages.iloc[i,j].split(',')]
                for t in lst_int:
                    TypeDict[Percentages.loc[i,'Type']][Type].append(t/100)
            else:
                lst_int=Percentages.iloc[i,j]
                for t in range(SumDict[Type]):
                    TypeDict[Percentages.loc[i,'Type']][Type].append(lst_int/SumDict[Type]/100)
    Final=pd.DataFrame()
    Final['Name']=Students.Name
    Final['ID']=Students.ID
    Final['Email']=Students.Email
    for key in SumDict:
        Final[key]=0
    Final['Total']=0
    Final['Grade']='NA'

    for i in range(len(Final)):
        Type=Students.loc[i,'Type']
        avgc=0
        for j in range(len(SumDict)):
            key=list(SumDict.keys())[j]
            avg=[]       
            for k in range(SumDict[key]):
                if SumDict[key]>1:
                    column=key+' '+str(k+1)
                else:
                    column=key
                avg.append(Students.loc[i,column])
            rem=Percentages.loc[len(Percentages)-1,key]
            if rem>0:
                avg.remove(np.sort(avg)[:rem])
            avgt=np.dot(avg,TypeDict[Type][key][rem:])/(np.sum(TypeDict[Type][key][rem:])+10**-17)
            avgc=avgc+np.dot(avg,TypeDict[Type][key][rem:])*len(TypeDict[Type][key])/len(TypeDict[Type][key][rem:])
            Final.loc[i,key]=avgt
        Final.loc[i,'Total']=avgc
        for l in Letters.columns:
            if avgc>=Letters.loc[0,l]:
                Final.loc[i,'Grade']=l
                break 
    return Final    