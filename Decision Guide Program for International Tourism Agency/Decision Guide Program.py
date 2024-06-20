import pandas as pd
import os
from tkinter import filedialog as fd
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

def main():
    #check existing file, if exist, delete the old file
    if os.path.exists('international_tourism_data.xlsx') == True:
        os.remove('international_tourism_data.xlsx')
    
    #create workbook using xlsxwriter
    writer = pd.ExcelWriter('international_tourism_data.xlsx', engine='xlsxwriter')
    workbook = writer.book
    
    #create and add new worksheet to the workbook
    worksheet = workbook.add_worksheet('Overview')
    
    #formatting
    bold = workbook.add_format({'bold': True})
    
    #input file using tkinter for worksheet 'Overview'
    filename = fd.askopenfilename()
    
    pd_file = pd.read_excel(filename, 'Overview', skiprows=9, nrows=1, usecols='C')
    source = []
    source.extend(pd_file)
    
    pd_conducted = pd.read_excel(filename, 'Overview', skiprows=10, nrows=1, usecols='C')
    conducted = []
    conducted.extend(pd_conducted)
    
    pd_period = pd.read_excel(filename, 'Overview', skiprows=11, nrows=1, usecols='C')
    period = []
    period.extend(pd_period)
    
    pd_region = pd.read_excel(filename, 'Overview', skiprows=12, nrows=1, usecols='C')
    region = []
    region.extend(pd_region)
    
    pd_published = pd.read_excel(filename, 'Overview', skiprows=21, nrows=1, usecols='C')
    published = []
    published.extend(pd_published)
    
    pd_date = pd.read_excel(filename, 'Overview', skiprows=22, nrows=1, usecols='C')
    date = []
    date.extend(pd_date)
    
    pd_original = pd.read_excel(filename, 'Overview', skiprows=23, nrows=1, usecols='C')
    original = []
    original.extend(pd_original)
    
    pd_description = pd.read_excel(filename, 'Overview', skiprows=8, nrows=1, usecols='E')
    description = []
    description.extend(pd_description)
    
    worksheet.write('A1', 'Source', bold)
    worksheet.write('A2', 'Conducted By', bold)
    worksheet.write('A3', 'Survey Period', bold)
    worksheet.write('A4', 'Region', bold)
    
    worksheet.write('A6', 'Published By', bold)
    worksheet.write('A7', 'Publication Date', bold)
    worksheet.write('A8', 'Original Source', bold)
    worksheet.write('A9', 'Description', bold)
    
    worksheet.write('B1', source[0])
    worksheet.write('B2', conducted[0])
    worksheet.write('B3', period[0])
    worksheet.write('B4', region[0])
    
    worksheet.write('B6', published[0])
    worksheet.write('B7', date[0])
    worksheet.write('B8', original[0])
    worksheet.write('B9', description[0])
    
    worksheet.set_default_row(20)
    worksheet.set_column(0, 0, 30)
    worksheet.set_column(1, 1, 100)
    
    #write the grabbed data to a new worksheet named 'Data'
    df_data = pd.read_excel(filename, 'Data', skiprows=4)
    df_data.to_excel(writer, sheet_name='Data', index=False)
    
    worksheet = writer.sheets['Data']
    
    worksheet.set_default_row(20)
    worksheet.set_column(0, 25, 30)
    
    dataAmerica_Mean = df_data['Americas'].mean()
    dataAmerica_Min = df_data['Americas'].min()
    dataAmerica_Max = df_data['Americas'].max()
    
    dataAsiaPacific_Mean = df_data['Asia Pacific'].mean()
    dataAsiaPacific_Min = df_data['Asia Pacific'].min()
    dataAsiaPacific_Max = df_data['Asia Pacific'].max()
    
    dataMiddleEast_Mean = df_data['Middle East'].mean()
    dataMiddleEast_Min = df_data['Middle East'].min()
    dataMiddleEast_Max = df_data['Middle East'].max()
    
    dataAfrica_Mean = df_data['Africa'].mean()
    dataAfrica_Min = df_data['Africa'].min()
    dataAfrica_Max = df_data['Africa'].max()
    
    dataEurope_Mean = df_data['Europe'].mean()
    dataEurope_Min = df_data['Europe'].min()
    dataEurope_Max = df_data['Europe'].max()
    
    #write the grabbed data into a new worksheet named 'Statistics'
    worksheet = workbook.add_worksheet('Statistics')
    
    worksheet.write('A1', 'Statistics by Region', bold)
    worksheet.write('A2', 'Mean', bold)
    worksheet.write('A3', 'Min', bold)
    worksheet.write('A4', 'Max', bold)
    
    worksheet.write('B1', 'Americas', bold)
    worksheet.write('B2', dataAmerica_Mean)
    worksheet.write('B3', dataAmerica_Min)
    worksheet.write('B4', dataAmerica_Max)
    
    worksheet.write('C1', 'Asia Pacific', bold)
    worksheet.write('C2', dataAsiaPacific_Mean)
    worksheet.write('C3', dataAsiaPacific_Min)
    worksheet.write('C4', dataAsiaPacific_Max)
    
    worksheet.write('D1', 'Middle East', bold)
    worksheet.write('D2', dataMiddleEast_Mean)
    worksheet.write('D3', dataMiddleEast_Min)
    worksheet.write('D4', dataMiddleEast_Max)
    
    worksheet.write('E1', 'Africa', bold)
    worksheet.write('E2', dataAfrica_Mean)
    worksheet.write('E3', dataAfrica_Min)
    worksheet.write('E4', dataAfrica_Max)
    
    worksheet.write('F1', 'Europe', bold)
    worksheet.write('F2', dataEurope_Mean)
    worksheet.write('F3', dataEurope_Min)
    worksheet.write('F4', dataEurope_Max)
    
    worksheet.set_default_row(20)
    worksheet.set_column(0, 25, 30)
    
    writer.save()
    
if __name__ == '__main__':
    main()