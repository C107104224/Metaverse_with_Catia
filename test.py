# import pymssql
# temp = ('HI',"20000")
#
# conn = pymssql.connect(
#     server='192.168.0.168',
#     user='sa', password='mbil2235',
#     database='test')
# cursor = conn.cursor(as_dict=True)
# # current_Date = datetime.now()
# # formatted_date = current_Date.strftime('%Y-%m-%d %H:%M:%S')
# # cursor.execute(
# #     """INSERT INTO [catia].[dbo].[CATIA_Part] ([Path],[FileName],[DateTime],[Length],[Width],[Height]) VALUES(%s%s%s%f%f%f)
# #     %(temp[0], temp[1], temp[2], temp[3], temp[4], temp[5])""")
# cursor.execute("INSERT INTO [dbo].[Store_Information] ([Store_Name],[Sales]) VALUES(%s,%s)", temp)
# conn.commit()
# print(cursor.rowcount, "Record inserted successfully into CATIA_Part table")
# cursor.close()
a = [1,2,3,4]
print(tuple(a))