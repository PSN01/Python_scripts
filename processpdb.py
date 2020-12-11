def processpdb():
    df5=pd.read_excel('dipole_rotation_C12_average_10_100.xlsx')
    t=df5.iloc[:,9].values.tolist()
    from biopandas.pdb import PandasPdb
    outputfile = "nvt_edit.pdb"
    outfile = open(outputfile,'w')
    ppdb = PandasPdb()
    ppdb.read_pdb('nvt.pdb')
    df1=ppdb.df['ATOM']
    df1.to_excel('nvtout.xlsx')
    df2 = df1.groupby([df1['residue_number'].ne(df1['residue_number'].shift()).cumsum(), 'residue_number']).size()
    df2.to_excel('nvtrange.xls')
    import numpy as np
    import pandas as pd
    x=df2.values.tolist()
    print(len(x))
    data=np.repeat(t, x)
    df3=pd.DataFrame(data)
    for i in df1.iterrows():
        df1.iloc[:, 15] = df3.iloc[:,0]
    ppdb.df['b_factor']=df1['b_factor']
    ppdb.to_pdb(outputfile, records=['ATOM'])

processpdb()
