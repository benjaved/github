import pandas as pd
try:    
    excel_file=r"C:\Users\MO40008729\Documents\pes1.xlsx"
    df=pd.read_excel(excel_file,sheetname='Raw data')

    ifile=open(r"C:\Users\MO40008729\Documents\inputconfig.property.txt")
    ver1="INPUT_COLUMN="
    ver1=ver1.lower()
    ver2="OUTPUT_COLUMN="
    ver2=ver2.lower()
    ver3="SEARCH_STRING="
    ver3=ver3.lower()
 
    s1=[]
    ic0=[]
    oc0=[]
    ss0=[]
    ver11=ver1.split("=")
    ver22=ver2.split("=")
    ver33=ver3.split("=")
    for i in range(0,100):
        s1=ifile.readline()
        s1=s1.replace('\n','')
        s1=s1.lower()
        slist=s1.split()
        for i in range(0,len(slist)):
            temp2=slist[i].split("=")
            if len(temp2)>1:
           
                        
                if temp2[0]==ver11[0]: 
                    ic0.append(temp2[i+1])
                if  temp2[0]==ver22[0]:
                    oc0.append(temp2[i+1])
                if temp2[0]==ver33[0]:
                    ss0.append(temp2[i+1])
            
        for l1 in range(0,len(slist)):
            temp=slist[l1]

            if temp.lower()==ver1:
                temp1=slist[l1+1]
                ic0.append(temp1)  
            
            elif temp.lower()==ver2:
                temp1=slist[l1+1]
                oc0.append(temp1)

            elif temp.lower()==ver3:
                temp1=slist[l1+1]
                ss0.append(temp1)
            
    
        
    ifile.close()
   
    ic0=[x for x in ic0 if x]
    ss0=[x for x in ss0 if x]
    oc0=[x for x in oc0 if x]
    
       
    ic=ic0[0].split(",")
    ss=ss0[0].split(",")
    oc=oc0[0].split(",")

    df=df.apply(lambda x: x.astype(str).str.lower())
    df.columns = map(str.lower, df.columns)
    df1={} 
    for v in oc:
        df1.update({v:df[v]})
        
        data1=pd.DataFrame.from_dict(df1)

    indx=[]
    len1=len(df[ic[0]])
    for i2 in range(0,len(ic)):
        scol=df[ic[i2]]    
        for i1 in range(0,len1):
            for ja in ss:
                if scol[i1].lower()==ja.lower():
                    indx.append(i1)
                else:    
                    pa=scol[i1].split()    
                    for xa in pa:
                        if ja.lower()==xa.lower():
                            indx.append(i1)
                        
    def Remove(duplicate):
             final_list = []
             for num in duplicate:
                 if num not in final_list:
                     final_list.append(num)
             return final_list
    sdata=[]
    indx=Remove(indx)
    for im in indx:
        sdata.append(data1.iloc[im])
        

    dfl=pd.DataFrame.from_dict(sdata)
    dfl=dfl.reset_index()
    dfl.columns = map(str.upper, dfl.columns)
    col=[]
    temp10=[]
    col1=[]
    col2=[]
    
    try:
         for i5 in range(0,len(oc)): 
             if oc[i5].lower()=="demand_id":
                 excel_file2=r"C:\Users\MO40008729\Documents\result.xlsx"
                 dfo=pd.read_excel(excel_file2,sheetname='Result Sheet')
                 tempo=dfo['DEMAND_ID']
                 tempn=dfl['DEMAND_ID']
             for i in range(0,len(tempn)):
                 set=0                    
                 for j in range (0,len(tempo)):
                     if tempn[i]==tempo[j]:
                         set=1
                         break
                 if set==0:
                     col.append(i)
                         
             for i in range(0,len(tempo)):
                 set=0                    
                 for j in range(0,len(tempn)):
                     if tempo[i]==tempn[j]:
                         set=1
                         break
                 if set==0:
                     col1.append(tempo[i])  
                     col2.append(i)
                
                    
                             
         dfm=pd.DataFrame.to_dict(dfl)
         list3=[]
         i=len(col)
         r=len(dfl['DEMAND_ID'])
         for x in range(0,r):
             list3.append("OLD")
           
         for x in range(0,len(dfl['DEMAND_ID'])):
             for b in range(0,i):
                 if x==col[b]:
                     list3[x]="NEW"
                     
                
         df4 = pd.DataFrame({'remark':list3})
         dfm.update({'REMARK':df4['remark']})
         dfm=pd.DataFrame.from_dict(dfm)  
         
         
         col2=Remove(col2)
         
         count=0 
          
         for i in col2:
             dfm=dfm.append(dfo.iloc[[i]])
             count+=1
             dfm=dfm.reset_index(drop=True)
      
      
         tes=dfm['REMARK']
         count=len(tes)-count
         for i in range(count,len(tes)):
            dfm.at[i,"REMARK"]="DELETED"
                
         writer=pd.ExcelWriter(r'C:\Users\MO40008729\Documents\result.stat.xlsx',engine='xlsxwriter')
         dfm.to_excel(writer,sheet_name="Result Sheet")
         writer.save()           
         writer=pd.ExcelWriter(r'C:\Users\MO40008729\Documents\result.xlsx',engine='xlsxwriter')
         dfl.to_excel(writer,sheet_name="Result Sheet")
         writer.save()
                    
    except:        
        writer=pd.ExcelWriter(r'C:\Users\MO40008729\Documents\result.xlsx',engine='xlsxwriter')
        dfl.to_excel(writer,sheet_name="Result Sheet")
        writer.save()
        print("done")
    
except:
        print("no file found or output/input coloumn is wrong")
        print("EVEN MAKE SURE THAT THE FILE IS CLOSE")





