import xlrd

book = xlrd.open_workbook('Приложение №2.xlsx')
sheet = book.sheet_by_index(1)
num_rows = sheet.nrows
num_col = sheet.ncols
r = ''

class a(object):
    def zentrovka_and_control_zagruzki(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'центровка и контроль загрузки' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) # рили или название функции
                    s =+ 1
                    #аааааааааа я не знаю, я хочу спать, но не могу уснуть,
                    #потому что обожрался ааааааааааааа памагити меня держат
                    #насильно
    def upravlenie_bezopasnostu_poletov(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'управление безопасностью полетов' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2)
                    s =+ 1
                    
    def opasnie_gruzi(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'опасные грузы' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def opasnie_gruzi_organizacii_possazirskix_perevozok(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'опасные грузы;организация поссажирских перевозок' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def aviacionna_bezoasnosnt(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'авиационная безопасность' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def aviaciono_spasatelnoe_obespechenie_poletov(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'авиационно-спасательное обеспечение полетов' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def protivoobledenitelnaa_zahita_vozduhhnix_sudov(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'противообледенительная защита воздушных судов' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def organizacia_nazemnogo_obsluzivania(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'организация наземного обслуживания' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def podgotovka_prepodovateley_auc(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'подготовка преподователей АУЦ' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1

    def organizacia_possazirskix_perevozok(self, name, programm, priorety, tema, grafic_raboti, grafic_smeni):

        global num_rows
        global sheet
        global num_col
        realy = [name, programm, priorety, tema, grafic_raboti, grafic_smeni]
        for i in range(num_rows):
            cell = sheet.cell(i, 3)
            if 'организация поссажирских перевозок' == cell:
                for j in range(num_col):
                    s = 0
                    cell2 = sheet.cell(i, j)
                    realy[s].append(cell2) 
                    s =+ 1
