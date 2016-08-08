import sys
import csv
import xlrd

chem_table = '\
CREATE TABLE IF NOT EXISTS `chem` (\n\
  `chem_id` int(11) NOT NULL AUTO_INCREMENT,\n\
  `casno` bigint(20) DEFAULT NULL,\n\
  `name` varchar(255) DEFAULT NULL,\n\
  PRIMARY KEY (`chem_id`)\n\
) ENGINE=MyISAM DEFAULT CHARSET=utf8 AUTO_INCREMENT=1;'

c_insert = 'INSERT INTO `chem` \
(`chem_id`,`casno`,`name`) VALUES '

properties_table = '\
CREATE TABLE IF NOT EXISTS `properties` (\n\
  `properties_id` int(11) NOT NULL AUTO_INCREMENT,\n\
  `chem_id` int(11) DEFAULT NULL,\n\
  `density` int(11) DEFAULT NULL,\n\
  `boilingpoint` int(11) DEFAULT NULL,\n\
  `freezingpoint` int(11) DEFAULT NULL,\n\
  PRIMARY KEY (`properties_id`)\n\
) ENGINE=MyISAM DEFAULT CHARSET=utf8 AUTO_INCREMENT=1;'

p_insert = 'INSERT INTO `properties` \
(`properties_id`,`chem_id`,`density`,`boilingpoint`,`freezingpoint`) VALUES '

END = '\n\n'

def main():

    if len(sys.argv) < 2:
        print('Filename needed')
        sys.exit(0)

    try:
        filename = sys.argv[1]
        chem = {}
        chem_temp = ''
        index = 1
        pindex = 1
        chem_id = 1
        sep = ''

        sqlout = open('EcoTTC.sql','w')
        sqlout.write(chem_table+END+properties_table+END+p_insert)

        if filename.endswith('.csv'):
            with open(filename,'r') as csvfile:
                datareader = csv.reader(csvfile)
                for row in datareader:
                    key = row[1]+row[2]
                    if key not in chem:
                        chem[key] = index
                        chem_temp += sep+'\n('+str(index)+','+str(row[2])+',\''+str(row[1])+'\')'
                        chem_id = index
                        index += 1
                    else:
                        chem_id = chem[key]

                    sqlout.write(sep+'\n('+str(pindex)+','+str(chem_id)+','+row[3]+','+row[4]+','+row[5]+')')
                    sep = ','
                    pindex += 1
                sqlout.write(';'+END+c_insert+chem_temp+';')

        elif filename.endswith('.xls') or filename.endswith('xlsx'):
            with xlrd.open_workbook(filename) as workbook:
                sheet_names = workbook.sheet_names()
                sheet = workbook.sheet_by_name(sheet_names[0])
                for rowx in range(1,sheet.nrows):
                    row = []
                    for colx in range(sheet.ncols):
                        row.append(sheet.cell_value(rowx,colx))
                    key = row[1]+str(row[2])
                    if key not in chem:
                        chem[key] = index
                        chem_temp += sep+'\n('+str(index)+','+str(row[2])+',\''+str(row[1])+'\')'
                        chem_id = index
                        index += 1
                    else:
                        chem_id = chem[key]

                    sqlout.write(sep+'\n('+str(pindex)+','+str(chem_id)+','+row[3]+','+row[4]+','+row[5]+')')
                    sep = ','
                    pindex += 1
                sqlout.write(';'+END+c_insert+chem_temp+';')


        else:
            print('Incompatible file type!')

        sqlout.close()

    except IOError as e:
        print('Unable to open file')

main()
