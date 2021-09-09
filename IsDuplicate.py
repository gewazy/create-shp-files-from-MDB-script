import pyodbc, time



# łączenie z bazą
conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=c:\PL-182_Husow_3D\!_PL182PY\baza\PL-182.mdb;')

cnxn = pyodbc.connect(conn_str)  # łączenie z bazą danych
crsr = cnxn.cursor()
cnxn.autocommit = True

# Zerowanie isDuplicate
print('zeruje')
crsr.execute("UPDATE [POSTPLOT] SET [POSTPLOT].`IsDuplicate` = NULL ")



start = time.perf_counter()
'''
# wypełnianie isDuplicate
wszystko = [item[0] for item in crsr.execute('SELECT [POSTPLOT].`Station (text)` FROM [POSTPLOT] where ([POSTPLOT].`Track` Between 1175 And 1930 or [POSTPLOT].`Track` Between 4060 And 4550) GROUP BY [POSTPLOT].`Station (text)` HAVING COUNT ([POSTPLOT].`Station (text)`) > 1').fetchall()]


#print(wszystko)
print('wypeł')
start = time.perf_counter()

for i in wszystko:
    crsr.execute(f"UPDATE [POSTPLOT] SET [POSTPLOT].`IsDuplicate` = '1' where [POSTPLOT].`Station (text)` like '{i}'")
'''
print('wypeł')
#crsr.execute(f"UPDATE [POSTPLOT] SET [POSTPLOT].`IsDuplicate` = 1 WHERE [POSTPLOT].`Station (text)` = `Station (text)` from (SELECT [POSTPLOT].`Station (text)`  FROM [POSTPLOT]  GROUP BY [POSTPLOT].`Station (text)` HAVING COUNT ([POSTPLOT].`Station (text)`) > 1 )")

upd = 'UPDATE ' \
      '[POSTPLOT], ' \
      'SET' \
      '[POSTPLOT].`IsDuplicate` = NULL ' \
      'where exists (SELECT [POSTPLOT].`Station (text)` FROM [POSTPLOT]  GROUP BY [POSTPLOT].`Station (text)` HAVING COUNT ([POSTPLOT].`Station (text)`) = 1 )'



crsr.execute(upd)


print('koniec')

end = time.perf_counter()
run_time = end -start
print(f'Wypełnienie IsDuplicate wykonano w: {run_time:.4f} sekund\n')


cnxn.close()
