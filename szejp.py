import pyodbc
import shapefile
import time

start = time.perf_counter()

# pola do shp:
pola_post = ["w.field('Stationtex', 'C', size=16)", "w.field('Stationval', 'N', size=16, decimal=3)",
             "w.field('Track', 'N', size=16)", "w.field('Bin', 'N', size=16)", "w.field('Descriptor', 'C', size=32)",
             "w.field('SurveyMode', 'C', size=16)", "w.field('SurveyMod0', 'C', size=5)",
             "w.field('HI', 'N', size=16, decimal=3)", "w.field('OffsetInli', 'N', size=16, decimal=3)",
             "w.field('OffsetCros', 'N', size=16, decimal=3)", "w.field('NumberofSa', 'C', size=5)",
             "w.field('PDOP', 'N', size=16,decimal=3)", "w.field('JulianDate', 'N', size=16, decimal=0)",
             "w.field('SurveyTime', 'C', size=8)", "w.field('SurveyTim1', 'C',size=8)",
             "w.field('Comment', 'C', size=32)", "w.field('DownloadFi', 'C', size=32)",
             "w.field('ReceiverTy', 'C', size=32)", "w.field('ReceiverSN', 'C', size=32)",
             "w.field('NumberOfEp', 'N', size=16, decimal=3)", "w.field('Occupation', 'N', size=16,decimal=3)",
             "w.field('AcquiredJu', 'N', size=16,decimal=0)", "w.field('Indeks', 'N', size=8, decimal=0)",
             "w.field('Status', 'N', size=8, decimal=0)", "w.field('IsDuplicat', 'N', size=8,decimal=0)",
             "w.field('Surveyor', 'C', size=16)", "w.field('Descriptio', 'C', size=32)",
             "w.field('Descripti2', 'C', size=32)", "w.field('depth', 'N', size=16, decimal=3)",
             "w.field('drname', 'C', size=16)", "w.field('dreq', 'C', size=16)", "w.field('drdate', 'C', size=16)",
             "w.field('UwagiBiuro', 'C', size=64)", "w.field('PPV', 'C', size=10)",
             "w.field('IsNNjoin', 'N', size=8,decimal=0)"]

pola_pre = ["w.field('Stationtex', 'C', size=16)", "w.field('Stationval', 'N', size=16, decimal=3)",
            " w.field('Track', 'N', size=16)", "w.field('Bin', 'N', size=16)"]

pola_post_kr = ["w.field('Stationtex', 'C', size=16)", "w.field('Stationval', 'N', size=16, decimal=3)",
                "w.field('Track', 'N', size=16)", "w.field('Bin', 'N', size=16)", "w.field('Descriptor', 'C', size=32)",
                "w.field('JulianDate', 'N', size=16, decimal=0)", "w.field('Comment', 'C', size=32)",
                "w.field('Indeks', 'N', size=8, decimal=0)", "w.field('Status', 'N', size=8, decimal=0)",
                "w.field('IsDuplicat', 'N', size=8,decimal=0)", "w.field('Descriptio', 'C', size=32)",
                "w.field('Descripti2', 'C', size=32)", "w.field('UwagiBiuro', 'C', size=64)"]
pola_qc_domiar_s = ["w.field('Stationtex', 'C', size=16)",  "w.field('Stationval', 'N', size=16, decimal=3)",
               "w.field('Track', 'N', size=16)", "w.field('Bin', 'N', size=16)", "w.field('UwagiBiuro', 'C', size=64)",
               "w.field('Descriptor', 'C', size=32)", "w.field('SurveyMode', 'C', size=16)",
               "w.field('Status', 'N', size=8, decimal=0)", "w.field('Indeks', 'N', size=8, decimal=0)"]

pola_otg_t = ["w.field('id', 'N', size=32, decimal=3)", "w.field('Nazwa', 'C', size=16)"]


# Zapytania do bazy:
postplot_s = "Select [POSTPLOT].* From [POSTPLOT] " \
             "Where [POSTPLOT].`Station (value)` > 0 And [POSTPLOT].`Status` >= 0 And [POSTPLOT].`Track` Between 4060 And 4550"

postplot_r = "Select [POSTPLOT].* From [POSTPLOT] " \
             "Where [POSTPLOT].`Station (value)` > 0 And [POSTPLOT].`Status` >= 0 And [POSTPLOT].`Track` Between 1175 And 1930"

preplot_s = "Select [PREPLOT].* From [PREPLOT] " \
            "Where ([PREPLOT].`Track` Like '4*') And [PREPLOT].`IsSurveyed` = 0"

preplot_r = "Select [PREPLOT].* From [PREPLOT] " \
            "Where ([PREPLOT].`Track` Like '1*')  And [PREPLOT].`IsSurveyed` = 0 And[PREPLOT].`Bin` <> 0"

skipy_s = "Select [POSTPLOT].* From [POSTPLOT] " \
          "Where [POSTPLOT].`Status` = 0 And [POSTPLOT].`Station (value)` > 0 And [POSTPLOT].`Track` Between 4060 And 4550"


skipy_r = "Select [POSTPLOT].* From [POSTPLOT] " \
          "Where [POSTPLOT].`Status` = 0 And [POSTPLOT].`Station (value)` > 0 And [POSTPLOT].`Track` Between 1175 And 1930"

qc_domiar_s = "Select " \
              "[POSTPLOT].`Station (text)`," \
              "[POSTPLOT].`Station (value)`, " \
              "[POSTPLOT].`Track`, [POSTPLOT].`Bin`, " \
              "IIF ([POSTPLOT].`Status` in (3, 4, 5),[POSTPLOT].`COG Local Easting`,[POSTPLOT].`Local Easting`)," \
              "IIF ([POSTPLOT].`Status` in (3, 4, 5),[POSTPLOT].`COG Local Northing`,[POSTPLOT].`Local Northing`)," \
              "[POSTPLOT].`Uwagi_Biuro`, " \
              "[POSTPLOT].`Descriptor`, " \
              "[POSTPLOT].`Survey Mode (value)`," \
              "[POSTPLOT].`Status`," \
              "[POSTPLOT].`Indeks` " \
              "From [POSTPLOT] " \
              "Where " \
              "[POSTPLOT].`Station (value)` > 0 And [POSTPLOT].`Track` Between 4060 And 4550  And ( ([POSTPLOT].`Status` IN (2,4) And  [POSTPLOT].`Survey Mode (value)` Not In (3,5,6))  Or ( [POSTPLOT].`Status` = 5  And  [POSTPLOT].`Survey Mode (value)` In (3) And ([POSTPLOT].`Number of Satellites` < 5 Or [POSTPLOT].`PDOP` > 6 Or [POSTPLOT].`CQ` > 0.3) )  or [POSTPLOT].`Status` = 5 or [POSTPLOT].`Status` = 6 )"

qc_domiar_r = "Select [POSTPLOT].* From [POSTPLOT] " \
              "Where  [POSTPLOT].`Station (value)` > 0 And [POSTPLOT].`Track` Between 1175 And 1930  And [POSTPLOT].`Status` >=1 And [POSTPLOT].`Status` <= 11 And (( [POSTPLOT].`Survey Mode (value)` Not In (3,5,6) ) Or ( [POSTPLOT].`Survey Mode (value)` = 3 And ([POSTPLOT].`Number of Satellites` < 5 Or [POSTPLOT].`PDOP` > 6 Or [POSTPLOT].`CQ` > 0.3) ))  Order By [POSTPLOT].`Station (text)`"

uwagi_geodetow = "Select [POSTPLOT].* From [POSTPLOT] Where [POSTPLOT].`Station (text)` Like '@%'"

pom_studnie = "Select [POM_STUD].* From POM_STUD"

otg_teory = "Select [OTG_teory].* From [OTG_teory] Where ([OTG_teory].`IsSurveyed` not like '%')"

otg = "Select [OTG].* From [OTG] Where ([OTG].`Station (value)`>0)"



# łączenie z bazą
conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=..\01_database\PL-182 HUSOW.mdb;')
cnxn = pyodbc.connect(conn_str)  # łączenie z bazą danych
crsr = cnxn.cursor()

# tworzenie list z danymi
postplot_s = crsr.execute(postplot_s).fetchall()
postplot_r = crsr.execute(postplot_r).fetchall()
preplot_s = crsr.execute(preplot_s).fetchall()
preplot_r = crsr.execute(preplot_r).fetchall()
skipy_s = crsr.execute(skipy_s).fetchall()
skipy_r = crsr.execute(skipy_r).fetchall()
qc_domiar_r = crsr.execute(qc_domiar_r).fetchall()
qc_domiar_s = crsr.execute(qc_domiar_s).fetchall()
uwagi_geodetow = crsr.execute(uwagi_geodetow).fetchall()
pom_studnie = crsr.execute(pom_studnie).fetchall()
otg_teory = crsr.execute(otg_teory).fetchall()
otg = crsr.execute(otg).fetchall()

crsr.close()
cnxn.close()


# tworzenie szejfilow:
with shapefile.Writer('shp_files/Postplot (source)', shapeType=1) as w:
    for i in pola_post:
        eval(i)

    for item in postplot_s:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 SurveyMode=item[16], SurveyMod0=item[17], HI=item[18],  OffsetInli=item[23], OffsetCros=item[24],
                 NumberofSa=item[30], PDOP=item[31], JulianDate=item[34], SurveyTime=item[35], SurveyTim1=item[36],
                 Comment=item[45], DownloadFi=item[46], ReceiverTy=item[48], ReceiverSN=item[49], NumberOfEp=item[50],
                 Occupation=item[55], AcquiredJu=item[74], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Surveyor=item[79], Descriptio=item[80], Descripti2=item[81], depth=item[82], drname=item[83],
                 dreq=item[84], drdate=item[85], UwagiBiuro=item[87], PPV=item[88], IsNNjoin=item[90])


with shapefile.Writer('shp_files/Postplot (receiver)', shapeType=1) as w:
    for i in pola_post:
        eval(i)

    for item in postplot_r:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 SurveyMode=item[16], SurveyMod0=item[17], HI=item[18], OffsetInli=item[23], OffsetCros=item[24],
                 NumberofSa=item[30], PDOP=item[31], JulianDate=item[34], SurveyTime=item[35], SurveyTim1=item[36],
                 Comment=item[45], DownloadFi=item[46], ReceiverTy=item[48], ReceiverSN=item[49], NumberOfEp=item[50],
                 Occupation=item[55], AcquiredJu=item[74], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Surveyor=item[79], Descriptio=item[80], Descripti2=item[81], depth=item[82], drname=item[83],
                 dreq=item[84], drdate=item[85], UwagiBiuro=item[87], PPV=item[88], IsNNjoin=item[90])


with shapefile.Writer('shp_files/Preplot (receiver) do tyczenia', shapeType=1) as w:
    for i in pola_pre:
        eval(i)

    for item in preplot_r:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3])


with shapefile.Writer('shp_files/Preplot (source) do tyczenia', shapeType=1) as w:
    for i in pola_pre:
        eval(i)

    for item in preplot_s:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3])


with shapefile.Writer('shp_files/SKIP S', shapeType=1) as w:
    for i in pola_post_kr:
        eval(i)

    for item in skipy_s:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 JulianDate=item[34], Comment=item[45], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Descriptio=item[80], Descripti2=item[81], UwagiBiuro=item[87])

with shapefile.Writer('shp_files/SKIP R', shapeType=1) as w:
    for i in pola_post_kr:
        eval(i)

    for item in skipy_r:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 JulianDate=item[34], Comment=item[45], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Descriptio=item[80], Descripti2=item[81], UwagiBiuro=item[87])

with shapefile.Writer('shp_files/QC_DOMIAR_RP', shapeType=1) as w:
    for i in pola_post_kr:
        eval(i)

    for item in qc_domiar_r:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 JulianDate=item[34], Comment=item[45], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Descriptio=item[80], Descripti2=item[81], UwagiBiuro=item[87])


with shapefile.Writer('shp_files/QC_DOMIAR_VP', shapeType=1) as w:
    for i in pola_qc_domiar_s:
        eval(i)

    for item in qc_domiar_s:
        w.point(item[4], item[5])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], UwagiBiuro=item[6],
                 Descriptor=item[7], SurveyMode=item[8], Status=item[9], Indeks=item[10])


with shapefile.Writer('shp_files/Uwagi_geodetow', shapeType=1) as w:
    for i in pola_post_kr:
        eval(i)

    for item in uwagi_geodetow:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 JulianDate=item[34], Comment=item[45], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Descriptio=item[80], Descripti2=item[81], UwagiBiuro=item[87])

with shapefile.Writer('shp_files/pom_studnie', shapeType=1) as w:
    for i in pola_post_kr:
        eval(i)

    for item in pom_studnie:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 JulianDate=item[34], Comment=item[45], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Descriptio=item[80], Descripti2=item[81], UwagiBiuro=item[87])


with shapefile.Writer('shp_files/OTG_zrobione', shapeType=1) as w:
    for i in pola_post_kr:
        eval(i)

    for item in otg:
        w.point(item[9], item[10])
        w.record(Stationtex=item[0], Stationval=item[1], Track=item[2], Bin=item[3], Descriptor=item[4],
                 JulianDate=item[34], Comment=item[45], Indeks=item[75], Status=item[76], IsDuplicat=item[77],
                 Descriptio=item[80], Descripti2=item[81], UwagiBiuro=item[87])

with shapefile.Writer('shp_files/OTG_teory', shapeType=1) as w:
    for i in pola_otg_t:
        eval(i)

    for item in otg_teory:
        w.point(item[3], item[4])
        w.record(id=item[5], Nazwa=item[0])

end = time.perf_counter()
run_time = end -start

print(f'Wykonano pliki SHP w: {run_time:.4f} sekund\n')
