
import pyodbc
import pandas as pd
from datetime import datetime, timedelta
import configparser


def get_db_config():
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    db_config = {
        'server': config.get('database', 'server'),
        'database': config.get('database', 'database'),
        'username': config.get('database', 'username'),
        'password': config.get('database', 'password')
    }
    return db_config


def db_connection(server, database, uid, pwd):
    
    conn = pyodbc.connect(
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={server};'
        f'DATABASE={database};'
        f'UID={uid};'
        f'PWD={pwd}'
    )
    return conn

def extract_data_from_xls(): #file_path, sheet_name)

    # data = pd.read_excel(r'D:\Shivam]\\'+Date+'_'+Shift+'.xls', sheet_name=0)
    # data = pd.read_excel(Date+'_'+Shift+'.xls', sheet_name=0)
    data = pd.read_excel('20250225_A.xls', sheet_name=0)
    db_config = get_db_config()
    con = db_connection(db_config["server"], db_config["database"], db_config["username"], db_config["password"])
    cursor = con.cursor()

    now1 = datetime.now()
    yesterday = now1 - timedelta(days=1)
    now = str(now1.strftime("%H:%M:%S"))
    now2 = str(now1.strftime("%Y%m%d"))
    y1 = str(yesterday.strftime("%Y%m%d"))
    if (now > "00:00:00" and now < "07:59:59"):
        Shift = "C"
        Date = y1
    elif (now > "08:00:00" and now < "15:59:59"):
        Shift = "A"
        Date = now2
    elif (now > "16:00:00" and now < "23:59:59"):
        Shift = "B"
        Date = now2
    # data = pd.read_excel('20230807_A', sheet_name=0)
    #data = pd.read_excel('20230726_B.xls', sheet_name=0)
    Member_Name = data['Member_Name'].unique()
    for i in Member_Name:
        field1 = data[data["Member_Name"].str.contains(i,na=False) & (data['field10']=='field10_test')]['field1'].unique()
        for j in field1:
            df = data[data["Member_Name"].str.contains(i,na=False) & (data['field10']=='field10_test') & (data["field1"].str.contains(j,na=False))]
            target = data[data["Member_Name"].str.contains(i,na=False) & (data["field1"].str.contains(j,na=False))]['field6'].sum()
            field8 = data[data["Member_Name"].str.contains(i,na=False) & (data["field1"].str.contains(j,na=False))]['field8'].sum()
            
            actual = df['field6'].sum()
            value12 = target - actual
            value13 = field8/target * 100 if(target > 0) else 'NA'
            Shift_result = (df['Shift'].to_string(index=False))
            Database_Date_result = str(df['Database_Date'])
            Member_ID_result = (df['Member_ID'].to_string(index=False)[0:df['Member_ID'].to_string(index=False).find('\n')])
            field3_result = (df['field3'].to_string(index=False))
            field1_result = (df['field1'].to_string(index=False))[0]
            value11_result = (df['value11'].to_string(index=False))
            field4_result = (df['field4'].to_string(index=False))
            field10_result = (df['field10'].to_string(index=False))[0]
            field9_result = (df['field9'].to_string(index=False))[0]
            field2_result = (df['field2'].to_string(index=False))
            field5_result = (df['field5'].to_string(index=False))
            field7_result = (df['field7'].to_string(index=False))[0]
            field6_result = (df['field6']).sum()
            value11_result = (df['value11'].to_string(index=False))[0]
            field8_result = (df['field8'].to_string(index=False))

            
            cursor.execute("SET ANSI_WARNINGS OFF;")
            
            cursor.execute('''
                            INSERT INTO [dbo].[Test_Excel]
                            ([Date]
                            ,[Member_ID]
                            ,[Member_Name]
                            ,[field1]
                            ,[field2]
                            ,[field3]
                            ,[field4]
                            ,[field5]
                            ,[field6]
                            ,[field7]
                            ,[field8]
                            ,[field9]
                            ,[field10]
                            ,[shift]
                            ,[value11]
                            ,[value12]
                            ,[value13]
                            ,[Crdatetime])
                            VALUES (getdate(), ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, getdate())
                            ''',
                            #(Database_Date_result),
                            str(Member_ID_result),
                            i,
                            j,
                            str(field2_result),
                            str(field3_result),
                            str(field4_result),
                            str(field5_result),
                            str(target),
                            int(field6_result),
                            str(field8),
                            int(field9_result),
                            str(field10_result),
                            Shift,
                            str(value11_result),
                            int(value12),
                            int(value13))
                    
            cursor.execute("SET ANSI_WARNINGS ON;")
            
    con.commit()
    con.close()
    print("Data stored in db Successfully.")

if __name__ == "__main__":
    extract_data_from_xls()
