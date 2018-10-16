import pandas as pd
import datetime
import simplekml
import timeit
import argparse
import os.path
import sqlite3

pd.options.mode.chained_assignment = None

def export_csv(dataframe, export_filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to CSV...')
  csv_data = dataframe.loc[:,['Location','Country','State_Province','Hole_ID','Original_ID','Date','Water_Depth','Lat','Long','Elevation','Position','Sample_Type','mblf_T','mblf_B','IGSN']]
  csv_data.loc[:, 'Date'] = csv_data.loc[:, 'Date'].apply(lambda x: x.replace('-','') if pd.notnull(x) else x)

  csv_data.to_csv(export_filename, encoding='utf-8-sig', index=False, float_format='%g')

  print('Collection exported to {} in {} seconds.\n'.format(export_filename, round(timeit.default_timer()-export_start_time,2)))


def export_html(dataframe, export_filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to HTML...')
  html_data = '<tr><td>' + dataframe.loc[:,'Location'] + '</td><td>'
  html_data += dataframe.loc[:,'Lat'].apply(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe.loc[:,'Long'].apply(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe.loc[:,'Elevation'].apply(lambda x: '' if pd.isnull(x) else str(round(x))) + '</td><td>'
  html_data += dataframe.loc[:,'IGSN'].apply(lambda x: '' if pd.isnull(x) else x) + '</td></tr>'

  with open(export_filename, 'w') as f:
    for r in html_data:
      f.write(r+'\n')

  print('Collection exported to {} in {} seconds.\n'.format(export_filename, round(timeit.default_timer()-export_start_time,2)))


def export_kml(dataframe, export_filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to KML...')

  kml_data = dataframe.loc[:,['Location','Long','Lat']]

  kml_data['GECF']= (dataframe['Hole_ID'].apply(lambda x: 'LacCoreID: ' + str(x) if pd.notnull(x) else '') + 
                     dataframe['Original_ID'].apply(lambda x: ' / FieldID: ' + str(x) if pd.notnull(x) else '') + 
                     dataframe['Date'].apply(lambda x: ' / Date: ' + str(x) if pd.notnull(x) else '') + 
                     dataframe['Water_Depth'].apply(lambda x: ' / Water Depth: ' + str(x) + 'm ' if pd.notnull(x) else '') +
                     dataframe[['mblf_T', 'mblf_B']].apply(lambda x: (' / Sediment Depth: ' + (str(x[0]) if pd.notnull(x[0]) else '?') + '-' + (str(x[1]) if pd.notnull(x[1]) else '?') + 'm') if (pd.notnull(x[0]) or pd.notnull(x[1])) else '', axis=1) + 
                     dataframe['Position'].apply(lambda x: ' / Position: ' + str(x) if pd.notnull(x) else '') + 
                     dataframe['IGSN'].apply(lambda x: ' / IGSN: ' + str(x) if pd.notnull(x) else '') + 
                     dataframe['Sample_Type'].apply(lambda x: ' / Sample Type: ' + str(x) if pd.notnull(x) else ''))

  kml = simplekml.Kml(name='LacCore/CSDCO Core Collection')
  style = simplekml.Style()
  style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png'
  style.iconstyle.scale = 1.2
  style.iconstyle.color = 'ff0000cc'

  for i, k in kml_data.iterrows():
    if pd.notnull(k[1]) and pd.notnull(k[2]):
      pnt = kml.newpoint(name=k[0], coords=[(k[1], k[2])], description=k[3])
      pnt.style = style

  kml.save(path=export_filename)

  print('Collection exported to {} in {} seconds.\n'.format(export_filename, round(timeit.default_timer()-export_start_time,2)))


def process_holes(holes_database, export_filename):
  print('Loading {}...'.format(holes_database))
  load_start_time = timeit.default_timer()

  conn = sqlite3.connect(holes_database)
  df = pd.read_sql_query("SELECT * FROM boreholes", conn)

  print('Loaded {} in {} seconds.\n'.format(holes_database, round(timeit.default_timer()-load_start_time,2)))

  df.loc[:, ['Lat', 'Long']] = df.loc[:,['Lat', 'Long']].apply(lambda x: round(x, 4))
  df.loc[:, 'Water_Depth'] = df.loc[:,'Water_Depth'].apply(lambda x: round(x, 2))

  export_csv(df, export_filename + '.csv')
  export_html(df, export_filename + '.txt')
  export_kml(df, export_filename + '.kml')

  print('Finished processing {} in {} seconds.\n'.format(holes_database, round(timeit.default_timer()-load_start_time,2)))


if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Filter and convert the LacCore Holes Excel file into different formats needed for publishing.')
  parser.add_argument('db_file', type=str, help='Name of CSDCO database file.')
  parser.add_argument('-d', action='store_true', help='Export files with the date in the filename (e.g., collection_YYYYMMDD.csv).')

  args = parser.parse_args()

  if os.path.isfile(args.db_file):
    filename = 'collection_' + datetime.datetime.now().strftime('%Y%m%d') if args.d else 'collection'
    process_holes(args.db_file, filename)
  else:
    print('ERROR: file \'{}\' does not exist.\n'.format(args.db_file))
