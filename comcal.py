import os
import re
import sys
import datetime
import pandas as pd

sys.argv.pop(0)
if (len(sys.argv) == 1 and re.match(r'^([\-\/]+h|\-\-*help)$', sys.argv[0])) or not (len(sys.argv) > 0 and len(sys.argv) <= 4):
  print("    Usage:")
  print("        comcal [function] [selected] [reg] [bros] [start] [end]\n")
  print("        where:\n")
  print("      function: 'max' to display a list of max absences per class and day")
  print("                'pro' to produce PDF certificates for all students")
  print("      selected: process only selected entries from the students file, e.g 0 or 1")
  print("           reg: registry number, e.g. 235")
  print("          bros: create sibling tab, e.g. 1 = yes, 0 = no")
  print("         start: optional date to start processing, e.g. 20220901")
  print("           end: optional date to finish processing, e.g. 20230630")
  sys.exit()

cwd = os.path.dirname(os.path.realpath(__file__))
datafolder = os.path.join(cwd, 'data')
outfolder = os.path.join(cwd, 'out')
studentfile = os.path.join(datafolder, 'students.xlsx')
absentfile = os.path.join(datafolder, 'absent.xlsx')
classfolder = os.path.join(datafolder, 'classes')
classfile = os.path.join(datafolder, 'Ώρες λειτουργίας τμημάτων.xlsx')
extrafile = os.path.join(datafolder, 'extras.xlsx')
outfile = os.path.join(outfolder, 'out.xlsx')
numbers = ['α.', 'β.', 'γ.', 'δ.', 'ε.']
cols = 'B:V'
stcols = ['Επώνυμο μαθητή','Όνομα μαθητή','Ημ/νία','Σύνολο απουσιών']
skiprows = 15
errors = {
    'dates':    [],
    'bigger':   [],
    'exdates':  [],
    'notfound': [],
}

def main(argv):
    f = str(sys.argv[0])
    sel = 0
    r = 0
    br = 1
    s = 0
    e = 0
    if len(sys.argv) > 1:
        sel = int(sys.argv[1])
    if len(sys.argv) > 2:
        r = int(sys.argv[2])
    if len(sys.argv) > 3:
        br = int(sys.argv[3])
    if len(sys.argv) > 4:
        s = str(sys.argv[4])
    if len(sys.argv) > 5:
        e = str(sys.argv[5])
    mo = re.search('^(\d{4})(\d{2})(\d{2})$', str(s))
    if s and not mo:
        s = 0
    elif mo:
        s = pd.DataFrame({'Ημ/νία': [mo.group(1)+'-'+mo.group(2)+'-'+mo.group(3)]})
        s['Ημ/νία'] = pd.to_datetime(s['Ημ/νία'], format='%Y-%m-%d')
    mo = re.search('^(\d{4})(\d{2})(\d{2})$', str(e))
    if e and not mo:
        e = 0
    elif mo:
        e = pd.DataFrame({'Ημ/νία': [mo.group(1)+'-'+mo.group(2)+'-'+mo.group(3)]})
        e['Ημ/νία'] = pd.to_datetime(e['Ημ/νία'], format='%Y-%m-%d')
    if not r or re.search('\D', str(r)):
        r = 1
    r = [int(r)]
    if re.search('\D', str(br)):
        br = 1

    exdata = pd.concat(pd.read_excel(extrafile, sheet_name=None), ignore_index=True)
    exdata['Ονοματεπώνυμο'] = exdata['Ονοματεπώνυμο'].replace('[ \-]+', '', regex=True)
    fixdates(exdata, s, e)

    cldata = pd.read_excel(classfile, sheet_name=None)
    for cl in cldata.keys():
        cldata[cl].drop(cldata[cl].index[cldata[cl]['Ώρες'].isnull()], inplace = True)
        fixdates(cldata[cl], s, e)

    studentsmov = pd.read_excel(studentfile, sheet_name='Sheet1', usecols=cols)
    if sel and f != 'max':
        studentsmov.drop(studentsmov.index[studentsmov['selected'].isnull()], inplace = True)
    studentsmov['Ονοματεπώνυμο'] = studentsmov['Ονοματεπώνυμο'].replace('[ \-]+', '', regex=True)
    studentsmov['Αδέλφια'] = studentsmov['Αδέλφια'].replace('[ \-]+', '', regex=True)
    studentsmov['Αρ. Πρωτ.'] = ""
    studentsmov['Ημ/νία'] = ""
    studentsmov['bros'] = ""
    studentsmov['total'] = ""

    absdata = pd.concat(pd.read_excel(absentfile, sheet_name=None, skiprows=skiprows, usecols=stcols), ignore_index=True)
    absdata.drop(absdata.index[absdata['Ημ/νία'].isnull()], inplace = True)
    absdata.drop(absdata.index[absdata['Ημ/νία'] == 'Ημ/νία'], inplace = True)
    if absdata.dtypes['Ημ/νία'] == 'object':
        absdata['Ημ/νία'] = pd.to_datetime(absdata['Ημ/νία'])
    if absdata.dtypes['Σύνολο απουσιών'] == 'object':
        absdata['Σύνολο απουσιών'] = pd.to_numeric(absdata['Σύνολο απουσιών'])
    absdata['Ονοματεπώνυμο'] = ""
    absdata['Τμήμα'] = ""
    name = ''
    surname = ''
    for index, row in absdata.iterrows():
        if pd.isnull(row['Όνομα μαθητή']):
            absdata.loc[index, 'Επώνυμο μαθητή'] = surname
            absdata.loc[index, 'Όνομα μαθητή'] = name
        else:
            surname = row['Επώνυμο μαθητή']
            name = row['Όνομα μαθητή']
        sn = re.sub('[ \-]+', '', f"{surname} {name}")
        absdata.loc[index, 'Ονοματεπώνυμο'] = sn
    for st in absdata['Ονοματεπώνυμο'].unique():
        temp = {}
        for i in absdata[absdata['Ονοματεπώνυμο'] == st].index.values:
            if not absdata.loc[i, 'Ημ/νία'] in temp.keys():
                temp[absdata.loc[i, 'Ημ/νία']] = ''
            else:
                absdata.drop(i, inplace = True)
    for index, row in studentsmov.iterrows():
        if not row['Ονοματεπώνυμο'] in absdata['Ονοματεπώνυμο'].unique():
            absdata.loc[absdata.tail(1).index[0]+1] = {
                'Επώνυμο μαθητή':   row['Επώνυμο μαθητή'],
                'Όνομα μαθητή':     row['Όνομα μαθητή'],
                'Ημ/νία':           pd.to_datetime(0),
                'Σύνολο απουσιών':  pd.to_numeric(0),
                'Ονοματεπώνυμο':    row['Ονοματεπώνυμο'],
                'Τμήμα':            row['Τμήμα'],
            }
        else:
            for i in absdata[absdata['Ονοματεπώνυμο'] == row['Ονοματεπώνυμο']].index.values:
                absdata.loc[i, 'Τμήμα'] = row['Τμήμα']
    for index, row in exdata.iterrows():
        if exdata.loc[index, 'Ονοματεπώνυμο'] in absdata['Ονοματεπώνυμο'].unique() and exdata.loc[index, 'Ονοματεπώνυμο'] in studentsmov['Ονοματεπώνυμο'].unique():
            found = 0
            cl = ''
            for j in absdata[absdata['Ονοματεπώνυμο'] == row['Ονοματεπώνυμο']].index.values:
                cl = absdata.loc[j, 'Τμήμα']
                d = row['Ημ/νία']
                dd = datetime.datetime.strptime(str(absdata.loc[j, 'Ημ/νία']), '%Y-%m-%d %H:%M:%S').strftime('%Y-%m-%d')
                if (d == dd):
                    absdata.loc[j, 'Σύνολο απουσιών'] = absdata.loc[j, 'Σύνολο απουσιών'].astype(int) + int(row['Ώρες'])
                    found = 1
            if not found:
                clindex = cldata[cl][cldata[cl]['date'] == row['date']].index.values[0]
                full = int(cldata[cl].loc[clindex, 'Ώρες'])
                if full == row['Ώρες']:
                    absdata.loc[absdata.tail(1).index[0]+1] = {
                        'Επώνυμο μαθητή':   row['Επώνυμο'],
                        'Όνομα μαθητή':     row['Όνομα'],
                        'Ημ/νία':           row['date'],
                        'Σύνολο απουσιών':  row['Ώρες'],
                        'Ονοματεπώνυμο':    row['Ονοματεπώνυμο'],
                        'Τμήμα':            cl,
                    }
                else:
                    errors['exdates'].append(f"Class: {cl}\nDate: "+cldata[cl].loc[clindex, 'Ημ/νία']+f"\nRegistry Periods: "+str(full)+f"\nStudent absent in extra for "+str(int(row['Ώρες']))+f" periods\nStudent: "+row['Επώνυμο']+' '+row['Όνομα'])
                    continue
        elif not sel:
            errors['notfound'].append(f"Class: "+row['Τμήμα']+"\nStudent: "+row['Επώνυμο']+' '+row['Όνομα'])
            continue
    fixdates(absdata, s, e)

    out = {}
    if f == 'max':
        showmax(cldata, absdata, out)
    else:
        writer = pd.ExcelWriter(outfile, engine = 'openpyxl')
        studentssta = studentsmov.copy()
        studentsmov.drop(studentsmov.index[studentsmov['Δημ. ενότητα'].isnull()], inplace = True)
        studentssta.drop(studentssta.index[studentssta['Δημ. ενότητα'].notnull()], inplace = True)
        studentsmov['Απόσταση σε μέτρα'] = studentsmov['Απόσταση σε μέτρα'].astype(int)
        process(cldata, studentsmov, absdata, out, writer, r, br, 'Moving')
        process(cldata, studentssta, absdata, out, writer, r, br,'Staying')
        writer.close()
        for key in sorted(errors.keys()):
            if len(errors[key]):
                if key == 'dates':
                    print("ERROR: Date not found in class calendar!")
                elif key == 'exdates':
                    print("ERROR: Extra date does not appear in registry!")
                elif key == 'bigger':
                    print("ERROR: Bigger value found!")
                elif key == 'notfound':
                    print("ERROR: Extra name not found in students!")
                print('==========================================')
                for o in errors[key]:
                    print(o)
                    print('------------------------------------------')
                print("")

def fixdates(df, s, e):
    df.drop(df.index[df['Ημ/νία'].isnull()], inplace = True)
    if isinstance(s, pd.DataFrame):
        df.drop(df.index[df['Ημ/νία'] < s.loc[0, 'Ημ/νία']], inplace = True)
    if isinstance(e, pd.DataFrame):
        df.drop(df.index[df['Ημ/νία'] > e.loc[0, 'Ημ/νία']], inplace = True)
    df['date'] = df['Ημ/νία']
    df['Ημ/νία'] = pd.to_datetime(df['Ημ/νία']).dt.strftime('%Y-%m-%d')

def showmax(cldata, absdata, out):
    out['max'] = {}
    out['chk'] = {}
    for cl in absdata['Τμήμα'].unique():
        out['max'][cl] = {}
        out['chk'][cl] = {}
        for i in absdata[absdata["Τμήμα"] == cl].index.values:
            d = absdata.loc[i, 'Ημ/νία']
            if not d in out['max'][cl].keys():
                out['max'][cl][d] = 0
            a = absdata.loc[i, 'Σύνολο απουσιών'].astype(int)
            if a > out['max'][cl][d]:
                out['max'][cl][d] = a
            taught = cldata[cl][cldata[cl]['Ημ/νία'] == d].index.values
            if not len(taught):
                continue
            taught = cldata[cl].loc[taught[0], 'Ώρες'].astype(int)
            if abs(taught - a) < 3 and abs(taught - a) > 0:
                if not d in out['chk'][cl].keys():
                    out['chk'][cl][d] = {}
                out['chk'][cl][d][absdata.loc[i, 'Επώνυμο μαθητή']+' '+absdata.loc[i, 'Όνομα μαθητή']] = [a, taught]
    print('Max absent hours per day')
    print('==========================================')
    for cl in sorted(out['max'].keys()):
        print(cl)
        for d in sorted(out['max'][cl].keys()):
            print(f"{d}\t"+str(out['max'][cl][d]))
        print("")
    print('Check absent hours per day')
    print('==========================================')
    print("Date\t\tStud\tClass\tName")
    print('==========================================')
    for cl in sorted(out['chk'].keys()):
        if not len(out['chk'][cl].keys()):
            continue
        print(cl)
        for d in sorted(out['chk'][cl].keys()):
            for n in sorted(out['chk'][cl][d].keys()):
                print(f"{d}\t"+str(out['chk'][cl][d][n][0])+"\t"+str(out['chk'][cl][d][n][1])+f"\t{n}")
        print("")

def process(cldata, students, absdata, out, writer, r, br, title):
    for st in students['Ονοματεπώνυμο'].unique():
        found = students[students['Ονοματεπώνυμο'] == st].index.values
        if not len(found):
            continue
        k = found[0].astype(int)
        cl = students.loc[k, 'Τμήμα']
        temp = cldata[cl].copy()
        temp.drop(temp.index[temp['date'] < students.loc[k, 'Έναρξη']], inplace = True)
        temp.drop(temp.index[temp['date'] > students.loc[k, 'Λήξη']], inplace = True)
        cldates = temp['Ημ/νία'].unique()
        stdates = []
        if not cl in out.keys():
            out[cl] = {}
        out[cl][st] = []
        for i in absdata[absdata["Ονοματεπώνυμο"] == st].index.values:
            d = absdata.loc[i, 'Ημ/νία']
            a = absdata.loc[i, 'Σύνολο απουσιών'].astype(int)
            if not d in cldates:
                errors['dates'].append(f"Class: {cl}\nDate: {d}\nStudent: "+students.loc[k, 'Επώνυμο μαθητή']+' '+students.loc[k, 'Όνομα μαθητή'])
                continue
            j = temp[temp['Ημ/νία'] == d].index.values[0].astype(int)
            b = temp.loc[j, 'Ώρες'].astype(int)
            if a == b:
                stdates.append(d)
            elif a > b:
                stdates.append(d)
                errors['bigger'].append(f"{cl}\n{d}\nStudent "+students.loc[k, 'Επώνυμο μαθητή']+' '+students.loc[k, 'Όνομα μαθητή']+f" absent for {a} periods\nPeriods taught: {b}")
                continue
        out[cl][st] = sorted(list(set(cldates).difference(stdates)))
        students.loc[k, 'Έναρξη'] = out[cl][st][0]
        students.loc[k, 'Λήξη'] = out[cl][st][-1]
        students.loc[k, 'Ημ/νία'] = "\n".join(out[cl][st])
        students.loc[k, 'total'] = len(out[cl][st])
    studentsbros = students.copy()
    if br:
        bros = {}
        for b in students['Αδέλφια']:
            if b in bros.keys():
                continue
            num = students[students['Αδέλφια'] == b].index.values
            if len(num) > 1:
                bros[b] = {}
                bros[b]['dates'] = {}
                bros[b]['names'] = []
                bros[b]['entry'] = ""
                counter = 0
                for i in num:
                    key = students.loc[i, 'Ονοματεπώνυμο']
                    cl = students.loc[i, 'Τμήμα']
                    if cl in out and key in out[cl].keys():
                        outnum = numbers[counter]
                        bros[b]['dates'][key] = out[cl][key]
                        bros[b]['names'].append(f"{outnum}\t"+students.loc[i, 'Επώνυμο μαθητή']+' '+students.loc[i, 'Όνομα μαθητή']+' του '+students.loc[i, 'Ονόματος πατέρα']+' και της '+students.loc[i, 'Ονόματος μητέρας']+' (ΑΜΜ '+students.loc[i, 'Αριθμός μητρώου'].astype(str)+', Τάξη '+students.loc[i, 'Τάξη']+')')
                        counter = counter + 1
                        if isinstance(bros[b]['entry'], str):
                            bros[b]['entry'] = students.loc[i]
                        else:
                            students.drop(students.index[students['Ονοματεπώνυμο'] == key], inplace = True)
                            studentsbros.drop(studentsbros.index[studentsbros['Ονοματεπώνυμο'] == key], inplace = True)
        for i in studentsbros.index.values:
            b = studentsbros.loc[i, 'Αδέλφια']
            if b in bros.keys():
                temp = []
                studentsbros.loc[i, 'bros'] = "\n".join(bros[b]['names'])
                for key in bros[b]['dates']:
                    temp = temp + bros[b]['dates'][key]
                temp = sorted(list(set(temp)))
                # temp = sorted(list(dict.fromkeys(temp)))
                studentsbros.loc[i, 'Ημ/νία'] = "\n".join(temp)
                studentsbros.loc[i, 'Έναρξη'] = temp[0]
                studentsbros.loc[i, 'Λήξη'] = temp[-1]
                studentsbros.loc[i, 'bros'] = "\n".join(bros[b]['names'])
                studentsbros.loc[i, 'total'] = len(temp)
        newstudents = students[~students['Αδέλφια'].isin(bros.keys())].copy()
        students = newstudents
        studentsbros = studentsbros[studentsbros['Αδέλφια'].isin(bros.keys())]
    for st in students['Ονοματεπώνυμο'].unique():
        found = students[students['Ονοματεπώνυμο'] == st].index.values
        students.loc[found[0].astype(int), 'Αρ. Πρωτ.'] = r[0]
        r[0] = int(r[0]) + 1
    students = students.astype(str)
    students['Έναρξη'] = students['Έναρξη'].replace("'", '', regex=True)
    students['Λήξη'] = students['Λήξη'].replace("'", '', regex=True)
    students.to_excel(writer, sheet_name = title+' Students')
    if br:
        for st in studentsbros['Ονοματεπώνυμο'].unique():
            found = studentsbros[studentsbros['Ονοματεπώνυμο'] == st].index.values
            studentsbros.loc[found[0].astype(int), 'Αρ. Πρωτ.'] = r[0]
            r[0] = r[0] + 1
        studentsbros['Έναρξη'] = pd.to_datetime(studentsbros['Έναρξη'], format='%Y-%m-%d')
        studentsbros['Λήξη'] = pd.to_datetime(studentsbros['Λήξη'], format='%Y-%m-%d')
        studentsbros = studentsbros.astype(str)
        studentsbros['Έναρξη'] = studentsbros['Έναρξη'].replace("'", '', regex=True)
        studentsbros['Λήξη'] = studentsbros['Λήξη'].replace("'", '', regex=True)
        studentsbros.to_excel(writer, sheet_name = title+' Siblings')

if __name__ == "__main__":
    main(sys.argv[0:])
    sys.exit()
