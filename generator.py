import os

import matplotlib.pyplot as plt  # v 2.2.2
import pandas as pd  # v 0.23.1
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle  # reportlab v 3.4.0

# XlsxWriter v 1.0.5
# numpy v 1.14.5

# load the CSV file to a DataFrame
raw = pd.read_csv("data_file.csv", sep=";", na_values=["UNK", "N/A"],
                  names=['img', 'plate', 'country', 'is_ok', 'plr', 'ctr', 'conf3', 'conf1', 'conf2'], skiprows=1)
'''
img = Image name
plate = Plate recorded
country = country/state recorded
is_ok = Is it expected to read based on picture's quality
plr = plate read
ctr = country/state read
conf3 = Confidence level combined
conf1 = Confidence level of Plate guessing
conf2 = Confidence level of Country/State guessing
'''
# Exclude images where GT == N/A and ARH == N/A
raw.dropna(inplace=True, how='all', subset=['plate', 'country', 'plr', 'ctr'])

###################################################################################
###  List of variables used  ######################################################
###################################################################################
'''
arh - list of DF's containing ARH evaluation for Text[0] State[1] and Combined[2]
gc - list of 3 template DF's sorted by gc levels
gc1 - list of 3 DF's prepared for sorted ARH evaluation
gc2 - list of 3 DF's prepared for cumulative sum
gc3 - list of 3 DF's prepared for inverted cumulative sum
writer - the stream to excel file via pandas
we cannot add charts only with pandas
workbook - XlsxWriter connection with excel file
worksheet - active sheet in an excel file
'''
###############################################################################################################
###############################################################################################################
###  Lists used ###############################################################################################
###############################################################################################################
###############################################################################################################

# index list
index = ["-Total", "Success", "-Fail",
         "GT!=NA&ARH!=NA",
         "GT=NA&ARH!=NA",
         "GT!=NA&ARH=NA"]
# querry list of 3 lists, each for Text State and Combined
ql = [
    [
        "plr == plr or plate == plate",  # at least one not nan                             -Total
        "plate == plr",  # ok, both not nan                                               -Success
        "plate != plr and (plate == plate or plr == plr)",  # not ok, at least one not nan-Fail
        "plate != plr and (plate == plate and plr == plr)",  # not ok, both not nan       GT!=NA&ARH!=NA
        "plate != plate and plr == plr",  # not ok, plate is nan                          GT!=NA&ARH!=NA
        "plate == plate and plr != plr",  # not ok, plr is nan                            GT!=NA&ARH!=NA
    ], [
        "ctr == ctr or country == country",  # at least one not nan                         -Total
        "country == ctr",  # ok, both not nan                                             -Success
        "country != ctr and (ctr == ctr or country == country)",  # not ok, at least one not nan -Fail
        "country != ctr and (country == country and ctr == ctr)",  # not ok, both not nan GT!=NA&ARH!=NA
        "country != country and ctr == ctr",  # not ok, country is nan                    GT!=NA&ARH!=NA
        "country == country and ctr != ctr",  # not ok, ctr is nan                        GT!=NA&ARH!=NA
    ], [
        "(plr == plr or plate == plate) and (ctr == ctr or country == country)",
        # at least one not nan plate and at least one not nan country
        "plate == plr and country == ctr",  # ok, both not nan
        "(plate != plr or country != ctr) and (ctr == ctr or country == country) and (plr == plr or plate == plate)",
        # at least one fail, at least one not nan plate and at least one not nan country
        "(plate != plr and country == ctr) and (ctr == ctr or country == country) and (plr == plr or plate == plate)",
        # fail text, at least one not nan plate and at least one not nan country
        "(plate == plr and country != ctr) and (ctr == ctr or country == country)",
        # fail state, at least one not nan plate and at least one not nan country
        "(plate != plr and country != ctr)  and (plr == plr or plate == plate) and (ctr == ctr or country == country)",
        # fail both
    ]]


################################################################################################################
###  List of functions used  ###################################################################################
################################################################################################################

def calc_sum(df, n=0):
    # df.query() returns new df, df.query()['img'] returns Series, df.query()['img'].count() returns value
    # Preforms query i times and appends the list by number of elements in the query
    # We use ['img'] instead of [ df.query(qlt[i].keys()[0] ] because its way faster
    suma = [df.query(ql[n][i])['img'].count() for i in range(6)]
    return suma


def calc_prc(df, n=0):  # used in ARH evaluation for text state both only
    mx = df.query(ql[n][0])['img'].size
    procent_of_total = [str(round(df.query(ql[n][i])['img'].size / mx * 100, 1)) + '%' for i in range(6)]
    return procent_of_total


# The conf list indicates the confidence level we are looking at by the n argument
confi = ['conf1', 'conf2', 'conf3']


def calc_min_gc(df, n=0):  # used in ARH evaluation for text state both only
    min_gc = [str(df.query(ql[n][i])[confi[n]].min()) + '%' for i in range(6)]
    return min_gc


def calc_max_gc(df, n=0):
    max_gc = [str(df.query(ql[n][i])[confi[n]].max()) + '%' for i in range(6)]
    return max_gc


def calc_avg_gc(df, n=0):
    avg_gc = [str(round(df.query(ql[n][i])[confi[n]].mean(), 1)) + '%' for i in range(6)]
    return avg_gc


def insert_perc(df):  # used in ARH evaluation by confidence levels
    s = df['-Total'][10]  # Colsum value

    df.insert(2, "%", [str(round(df['Success'][i] / s * 100, 1)) + '%' for i in range(df['Success'].size)],
              allow_duplicates=True)
    df.insert(4, "%", [str(round(df['-Fail'][i] / s * 100, 1)) + '%' for i in range(df['Success'].size)],
              allow_duplicates=True)
    df.insert(6, "%", [str(round(df['GT!=NA&ARH!=NA'][i] / s * 100, 1)) + '%' for i in range(df['Success'].size)],
              allow_duplicates=True)
    df.insert(8, "%", [str(round(df['GT=NA&ARH!=NA'][i] / s * 100, 1)) + '%' for i in range(df['Success'].size)],
              allow_duplicates=True)
    df.insert(10, "%", [str(round(df['GT!=NA&ARH=NA'][i] / s * 100, 1)) + '%' for i in range(df['Success'].size)],
              allow_duplicates=True)


def append_automation(df, n=0):  # used in inverse sum
    df['#Automation'] = df['-Total'] - df["GT!=NA&ARH=NA"]
    df['%Automation'] = [round(df['#Automation'][i] / df['-Total'][0] * 100, 1) if df['#Automation'][i] != 0 else 0 for
                         i in range(df['#Automation'].size)]

    if n == 0:
        df['#FP'] = df["GT=NA&ARH!=NA"] + df["GT!=NA&ARH!=NA"]
    else:
        df['#FP'] = df["-Fail"]
    df['%FP'] = [
        round(df['#FP'][i] / df['#Automation'][i] * 100, 1) if df['#FP'][i] != 0 and df['#Automation'][i] != 0 else 0
        for i in range(df['#FP'].size)]
    df[' '] = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, ' ']
    df['  '] = df['%FP']


def cdic(n=0):  # creates dict used to create DataFrame file (ARH evaluation)
    return {"Sum": calc_sum(raw, n),
            "%": calc_prc(raw, n),
            "Min GC": calc_min_gc(raw, n),
            "Max GC": calc_max_gc(raw, n),
            "Avg GC": calc_avg_gc(raw, n)
            }


def cdic2(n=0):  # creates dict used to create DataFrame file
    dic = {"[0-10)": calc_sum(raw.query("{}>=0 and {}<10".format(confi[n], confi[n])), n),
           "[10-20)": calc_sum(raw.query("{}>=10 and {}<20".format(confi[n], confi[n])), n),
           "[20-30)": calc_sum(raw.query("{}>=20 and {}<30".format(confi[n], confi[n])), n),
           "[30-40)": calc_sum(raw.query("{}>=30 and {}<40".format(confi[n], confi[n])), n),
           "[40-50)": calc_sum(raw.query("{}>=40 and {}<50".format(confi[n], confi[n])), n),
           "[50-60)": calc_sum(raw.query("{}>=50 and {}<60".format(confi[n], confi[n])), n),
           "[60-70)": calc_sum(raw.query("{}>=60 and {}<70".format(confi[n], confi[n])), n),
           "[70-80)": calc_sum(raw.query("{}>=70 and {}<80".format(confi[n], confi[n])), n),
           "[80-90)": calc_sum(raw.query("{}>=80 and {}<90".format(confi[n], confi[n])), n),
           "[90-100]": calc_sum(raw.query("{}>=90 and {}<=100".format(confi[n], confi[n])), n)
           }
    return dic


def add_czart(czart, n=0, chart_data=0):
    czart.add_series({
        'name': 'Success',
        'categories': "=Sheet1!$A${}:$A${}".format(chart_data, chart_data + 9),
        'values': "=Sheet1!$C${}:$C${}".format(chart_data, chart_data + 9),
        'fill': {'color': '#52ce33'},
    })
    if n == 0:
        # CUMSUM
        czart.add_series({
            'name': 'Fail',
            'categories': "=Sheet1!$A${}:$A${}".format(chart_data, chart_data + 9),
            'values': "=Sheet1!$D${}:$D${}".format(chart_data, chart_data + 9),
            'fill': {'color': 'red'},
        })
    else:
        # TOTAL
        czart.add_series({
            'name': 'Fail',
            'categories': "=Sheet1!$A${}:$A${}".format(chart_data, chart_data + 9),
            # '%' columns are added
            'values': "=Sheet1!$E${}:$E${}".format(chart_data, chart_data + 9),
            'fill': {'color': 'red'},
        })
    czart.set_x_axis({'name': 'Intervals'})
    czart.set_y_axis({'name': 'No. images'})


def insert_charts(czarts=[], row=0):  # puts charts in to the excel file
    czarts[0].set_size({'width': 550, 'height': 300})
    czarts[1].set_size({'width': 550, 'height': 300})
    czarts[2].set_size({'width': 550, 'height': 300})
    worksheet.insert_chart('A{}'.format(row), czarts[0], {'x_offset': 15, 'y_offset': 5})
    worksheet.insert_chart('H{}'.format(row), czarts[1], {'x_offset': 15, 'y_offset': 5})
    worksheet.insert_chart('E{}'.format(row + 16), czarts[2], {'x_offset': 15, 'y_offset': 5})


########################################################################################################################
########################################################################################################################
########################################################################################################################

# the stream to excel file via pandas
writer = pd.ExcelWriter('Report.xlsx', engine='xlsxwriter')

# ARH evaluation DataFrames initialization
arh = [pd.DataFrame(cdic(), index=index), pd.DataFrame(cdic(1), index=index), pd.DataFrame(cdic(2), index=index)]
####
# Template Dataframes sorted by GC levels
gc = [
    pd.DataFrame(cdic2(), index=index).transpose(),
    pd.DataFrame(cdic2(1), index=index).transpose(),
    pd.DataFrame(cdic2(2), index=index).transpose()
]
####
# ARH evaluation ordered by GC levels
gc1 = [pd.concat([gc[i], pd.DataFrame({'Total': [gc[i][j].sum() for j in gc[0].columns]}, index=index).transpose()]) for
       i in range(3)]
for i in range(3):
    insert_perc(gc1[i])
####
# Cumulative sum of images by the confidence level
gc2 = [gc[i].cumsum() for i in range(3)]
gc2 = [pd.concat([gc2[i], pd.DataFrame({'Total': [gc2[i][j].max() for j in gc[0].columns]}, index=index).transpose()])
       for i in range(3)]
for i in range(3):
    gc2[i]['-Total'] = gc1[i]['-Total'].copy()
####
# Global inverted cumulative sum
gc3 = [gc[i][::-1].cumsum()[::-1] for i in range(3)]
gc3 = [pd.concat([gc3[i], pd.DataFrame({'Total': [gc3[i][j].max() for j in gc[0].columns]}, index=index).transpose()])
       for i in range(3)]
for i in range(3):
    append_automation(gc3[i])
####

#############################################
### TO EXCEL
#############################################
# ARH
arh[0].to_excel(writer, startrow=5)
arh[1].to_excel(writer, startrow=13)
arh[2].to_excel(writer, startrow=21)

# ARH BY GC
gc1[0].to_excel(writer, startrow=31)
gc1[1].to_excel(writer, startrow=44)
gc1[2].to_excel(writer, startrow=57)

# CUM SUM
gc2[0].to_excel(writer, startrow=104)
gc2[1].to_excel(writer, startrow=117)
gc2[2].to_excel(writer, startrow=130)

# INV CUM SUM
gc3[0].to_excel(writer, startrow=177)
gc3[1].to_excel(writer, startrow=190)
gc3[2].to_excel(writer, startrow=203)

#####################################
## XlsxWriter Sheet initialization ##
#####################################
workbook = writer.book
worksheet = writer.sheets['Sheet1']

cell_format = workbook.add_format()
cell_format.set_align('right')

worksheet.set_column(0, 0, 18)
worksheet.set_column('E:M', 16)
worksheet.set_zoom(75)
##########################################
# formatting
##########################################
title = workbook.add_format({'bold': True, 'italic': True, 'font_size': 22})
subtitle = workbook.add_format({'bold': True, 'italic': True, 'font_size': 14})

worksheet.set_row(1, 30, title)
worksheet.set_row(3, 30, title)
worksheet.set_row(29, 30, title)
worksheet.set_row(102, 30, title)
worksheet.set_row(175, 30, title)
worksheet.write('A2', "Doesn’t include image where GT == N/A e ARH == N/A")
worksheet.write('A4', "ARH evaluation")
worksheet.write('A30', "ARH evaluation")
worksheet.write('A103', "Sum of images by the confidence level")
worksheet.write('A176', "Inverse sum of images by the confidence level")

worksheet.write('A6', "Text", subtitle)
worksheet.write('A14', "State", subtitle)
worksheet.write('A22', "Combined", subtitle)

worksheet.write('A32', "Text", subtitle)
worksheet.write('A45', "State", subtitle)
worksheet.write('A58', "Combined", subtitle)

worksheet.write('A105', "Text", subtitle)
worksheet.write('A118', "State", subtitle)
worksheet.write('A131', "Combined", subtitle)

worksheet.write('A178', "Text", subtitle)
worksheet.write('A191', "State", subtitle)
worksheet.write('A204', "Combined", subtitle)

#############################################
### CHARTS
#############################################
# list of chart objects
charts = [workbook.add_chart({'type': 'column', 'subtype': 'stacked'}) for i in range(9)]

# Filling charts with values
# chart object, 1 stands for table where '%' columns are added, row were to read values
add_czart(charts[0], 1, 33)
add_czart(charts[1], 1, 46)
add_czart(charts[2], 1, 59)

add_czart(charts[3], 0, 106)
add_czart(charts[4], 0, 119)
add_czart(charts[5], 0, 132)

add_czart(charts[6], 0, 179)
add_czart(charts[7], 0, 192)
add_czart(charts[8], 0, 205)

# unique name for each chart has to be set
charts[0].set_title({'name': 'Distribution of confidence levels TEXT'})
charts[1].set_title({'name': 'Distribution of confidence levels STATE'})
charts[2].set_title({'name': 'Distribution of confidence levels COMBINED'})
charts[3].set_title({'name': 'Sum of images along the scale of degrees\n of confidence for the TEXT'})
charts[4].set_title({'name': 'Sum of images along the scale of degrees\n of confidence for the STATE'})
charts[5].set_title({'name': 'Sum of images along the scale of degrees\n of confidence for the COMBINED'})
charts[6].set_title({'name': 'Inverse sum of images along the scale of degrees\n of confidence for the TEXT'})
charts[7].set_title({'name': 'Inverse sum of images along the scale of degrees\n of confidence for the STATE'})
charts[8].set_title({'name': 'Inverse sum of images along the scale of degrees\n of confidence for the COMBINED'})

# insert 3 charts at once at correct positions
insert_charts(charts[0:3], 70)
insert_charts(charts[3:6], 143)
insert_charts(charts[6:9], 216)

########################################################################################################################
#### Excel file is completed here ####
########################################################################################################################
writer.save()
########################################################################################################################
########################################################################################################################

plt.rcParams.update({'font.size': 7, 'axes.axisbelow': True})

c = canvas.Canvas('plik.pdf', pagesize=A4)

styleSmall = TableStyle([
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
    ('VALIGN', (0, 0), (0, -1), 'TOP'),
    ('VALIGN', (0, -1), (-1, -1), 'MIDDLE'),
    ('INNERGRID', (0, 0), (-1, -1), 0.10, colors.lightgrey),
    ('GRID', (0, 0), (-1, 0), 0.4, colors.grey),
    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
    ('BACKGROUND', (0, 0), (-1, 0), '#d3d3d3'),
    ('BACKGROUND', (0, 2), (-1, 2), '#CCFF90'),
    ('BACKGROUND', (0, 3), (-1, 3), '#ff8a80'),
])
styleLarge = TableStyle([
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
    ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
    ('VALIGN', (0, 0), (0, -1), 'TOP'),
    ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
    ('INNERGRID', (0, 0), (-1, -1), 0.10, colors.lightgrey),
    ('GRID', (0, 0), (-1, 0), 0.4, colors.grey),
    ('GRID', (0, 11), (-1, 11), 0.3, colors.grey),
    ('BOX', (0, 0), (-1, -1), 0.25, colors.black),
    ('BACKGROUND', (0, 0), (-1, 0), '#d3d3d3'),
    ('BACKGROUND', (0, 1), (-1, 1), '#eeeeee'),
    ('BACKGROUND', (0, 3), (-1, 3), '#eeeeee'),
    ('BACKGROUND', (0, 5), (-1, 5), '#eeeeee'),
    ('BACKGROUND', (0, 7), (-1, 7), '#eeeeee'),
    ('BACKGROUND', (0, 9), (-1, 9), '#eeeeee'),
    ('BACKGROUND', (0, 11), (-1, 11), '#E0E0E0'),
])


def create_pdf(arh, gc, inv, title='CARACTERES'):
    pwidth, height = A4
    ###############################################################################
    # Print Table #1
    ###############################################################################
    arh.insert(0, 'Type', arh.index.values)
    tabela = [arh.columns]
    for i in arh.index.values:
        tabela.append(list((arh.loc[i])))
    t = Table(tabela, [150, 67, 67, 67, 67, 67])  # == 485
    t.setStyle(styleSmall)
    w, h = t.wrap(pwidth, height)
    t.drawOn(c, 55, 640)
    ###############################################################################
    # Prepare Table #2
    ###############################################################################
    c.drawCentredString(pwidth / 2, 800, title)
    c.line(pwidth / 2 - c.stringWidth(title, "Helvetica", 12) / 2, 798,
           pwidth / 2 + c.stringWidth(title, "Helvetica", 12) / 2, 798)
    c.drawCentredString(pwidth / 2, 780, 'Resumo de Eesultados')
    c.drawCentredString(pwidth / 2, 590, 'Distribuição do Grau de Confiança')
    ################################################
    gct_pdf = gc[['-Total', 'Success', '-Fail', "GT!=NA&ARH!=NA", "GT=NA&ARH!=NA", "GT!=NA&ARH=NA"]].copy()
    gct_pdf.insert(0, 'GC\nLevels', gct_pdf.index.values)

    gct_pdf["GT!=NA&ARH!=NA"] = [str(i) + '%' for i in
                                 round(gct_pdf["GT!=NA&ARH!=NA"] / gct_pdf['-Total'][10] * 100, 1)]
    gct_pdf["GT=NA&ARH!=NA"] = [str(i) + '%' for i in round(gct_pdf["GT=NA&ARH!=NA"] / gct_pdf['-Total'][10] * 100, 1)]
    gct_pdf["GT!=NA&ARH=NA"] = [str(i) + '%' for i in round(gct_pdf["GT!=NA&ARH=NA"] / gct_pdf['-Total'][10] * 100, 1)]

    gct_pdf.rename(index=str, columns={"GT!=NA&ARH!=NA": "GT<>NA\nARH<>NA", "GT=NA&ARH!=NA": "GT=NA\nARH<>NA",
                                       "GT!=NA&ARH=NA": "GT<>NA\nARH=NA"}, inplace=True)
    gct_pdf.insert(3, '%', round(gct_pdf['Success'] / gct_pdf['-Total'][10] * 100, 1), allow_duplicates=True)
    gct_pdf.insert(5, ' %', round(gct_pdf['-Fail'] / gct_pdf['-Total'][10] * 100, 1), allow_duplicates=True)

    gct_pdf_wo_colsum = gct_pdf.drop('Total')

    ###############################################################################
    # Print Chart #1
    ###############################################################################

    plt.rcParams.update({'font.size': 8, 'axes.axisbelow': True})
    ind = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]  # the x locations for the groups
    width = 0.45  # the width of the bars: can also be len(x) sequence

    succ = tuple(gct_pdf_wo_colsum.iloc[:, 3])
    fail = tuple(gct_pdf_wo_colsum.iloc[:, 5])

    p1 = plt.bar(ind, succ, width, color='#CCFF90', edgecolor='black', linewidth=0.5)
    p2 = plt.bar(ind, fail, width, bottom=succ, color='#ff8a80', edgecolor='black', linewidth=0.5)

    plt.title('Distribuição dos graus de confinança para os {}'.format(title))
    plt.xticks(ind, tuple(gct_pdf_wo_colsum['GC\nLevels']))
    plt.yticks([0, 5, 10, 15, 20, 25, 30], ['0%', '5%', '10%', '15%', '20%', '25%', '30%'])
    plt.grid(axis='y')
    plt.legend((p1[0], p2[0]), ('Success', 'Fail'))

    if not os.path.exists('imgs'):
        os.makedirs('imgs')
    plt.savefig('imgs/chart1_{}.png'.format(title), dpi=800)
    c.drawImage('imgs/chart1_{}.png'.format(title), 20, 0, 555, 325)
    plt.clf()
    plt.close()
    ###############################################################################
    # Print Table #2
    ###############################################################################
    gct_pdf.iloc[:, 3] = [str(i) + '%' for i in gct_pdf.iloc[:, 3]]
    gct_pdf.iloc[:, 5] = [str(i) + '%' for i in gct_pdf.iloc[:, 5]]

    tabela = [gct_pdf.columns]
    for i in gct_pdf.index.values:
        tabela.append(list((gct_pdf.loc[i])))
    t = Table(tabela, [55])
    t.setStyle(styleLarge)
    w, h = t.wrap(pwidth, height)
    t.drawOn(c, 50, (height - 500))

    c.showPage()
    ###############################################################################
    # Prepare Table #3
    ###############################################################################
    c.setFont('Times-Bold', 14)
    c.drawCentredString(pwidth / 2, 760, 'Grau de Automação e Falsos Positivos')

    gct_inv_pdf = inv.copy()

    gct_inv_pdf.drop([' ', '  '], axis=1, inplace=True)
    gct_inv_pdf.drop('Total', inplace=True)

    gct_inv_pdf.insert(0, 'GC\nThreshold', [0, 10, 20, 30, 40, 50, 60, 70, 80, 90])

    gct_inv_pdf.rename(index=str, columns={"Success": 'Success\n(1)', "-Fail": 'Fail\n(2)'
        , "GT!=NA&ARH!=NA": "GT<>NA\nARH<>NA\n(3)", "GT=NA&ARH!=NA": "GT=NA\nARH<>NA\n(4)",
                                           "GT!=NA&ARH=NA": "GT<>NA\nARH=NA\n(5)",
                                           "#Automation": "#\nAutomati\non (6)", "%Automation": "%\nAutomati\non (7)",
                                           "#FP": "# FP\n(8)", "%FP": "% FP\n(9)"}, inplace=True)

    c.setFont("Helvetica", 9)
    c.drawString(90, 430, 'Notas:')
    c.drawString(160, 430, '#Automation(6) = (1)+(3)+(4)')
    c.drawString(160, 410, '%Automation(7) = #Automation(6)/Total')
    c.drawString(330, 430, '#FP(8) = (3)+(4)')
    c.drawString(330, 410, '%FP(9) = #FP(8)/#Automation(6)')

    ###############################################################################
    # Print Chart #2
    ###############################################################################
    width = 0.35  # the width of the bars: can also be len(x) sequence

    fig, ax = plt.subplots()

    automacio = tuple(gct_inv_pdf["%\nAutomati\non (7)"])
    rects1 = ax.bar([i - 0.03 for i in ind], automacio, width, color='#CCFF90', edgecolor='black', linewidth=0.5)

    fp = tuple(gct_inv_pdf["% FP\n(9)"])
    rects2 = ax.bar([i + width + 0.03 for i in ind], fp, width, color='#ff8a80', edgecolor='black', linewidth=0.5)

    ax.set_title('KPI Automação e FP - {}'.format(title))
    ax.set_xticks([(i + width / 2) for i in ind])
    ax.set_xticklabels(('0%', '10%', '20%', '30%', '40%', '50%', '60', '70', '80', '90'))
    ax.set_yticks([0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 108], ["0%", "10%", "20%", "30%", "40%", "50%",
                                                                      "60%", "70%", "80%", "90%", "100%"])
    ax.legend((rects1[0], rects2[0]), ('Automacio', 'Falsos Positivos'))

    # Adding value over bars
    for rect in rects1:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width() / 2., 1.01 * height, '%.1f' % height + str("%"), ha='center',
                va='bottom', fontsize=5.5)
    for rect in rects2:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width() / 2., 1.01 * height, '%.1f' % height + str("%"), ha='center',
                va='bottom', fontsize=5.5)

    plt.savefig('imgs/chart2_{}.png'.format(title), dpi=800)
    c.drawImage('imgs/chart2_{}.png'.format(title), 20, 0, 575, 400)
    plt.clf()
    plt.close()
    ###############################################################################
    # Print Table #3
    ###############################################################################
    gct_inv_pdf.iloc[:, 8] = [str(i) + '%' for i in gct_inv_pdf.iloc[:, 8]]
    gct_inv_pdf.iloc[:, 10] = [str(i) + '%' for i in gct_inv_pdf.iloc[:, 10]]

    tabela = [gct_inv_pdf.columns]
    for i in gct_inv_pdf.index.values:
        tabela.append(list((gct_inv_pdf.loc[i])))

    t = Table(tabela, [50], [40, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22])
    t.setStyle(styleLarge)
    w, h = t.wrap(pwidth, height)
    t.drawOn(c, 22, 460)
    c.showPage()


create_pdf(arh[0], gc1[0], gc3[0], 'CARACTERES')
create_pdf(arh[1], gc1[1], gc3[1], 'ESTADO')
create_pdf(arh[2], gc1[2], gc3[2], 'CARACTERES + ESTADO')

c.save()
