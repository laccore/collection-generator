import pandas as pd
import datetime
import simplekml
import timeit
import argparse
import os.path
import sqlite3

pd.options.mode.chained_assignment = None

def export_csv(dataframe, filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to CSV...')
  csv_data = dataframe[['Location','Country','State_Province','Hole_ID','Original_ID','Date','Water_Depth','Lat','Long','Elevation','Position','Sample_Type','mblf_T','mblf_B','IGSN']]
  csv_data['Date'] = csv_data['Date'].apply(lambda x: x.replace('-','') if pd.notnull(x) else x)

  csv_data.to_csv(filename, encoding='utf-8-sig', index=False, float_format='%g')

  print('Collection exported to {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-export_start_time,2)))


def export_html(dataframe, filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to HTML...')
  html_data = '<tr><td>' + dataframe['Location'] + '</td><td>'
  html_data += dataframe['Lat'].apply(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe['Long'].apply(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe['Elevation'].apply(lambda x: '' if pd.isnull(x) else str(round(x))) + '</td><td>'
  html_data += dataframe['IGSN'].apply(lambda x: '' if pd.isnull(x) else x) + '</td></tr>'

  with open(filename, 'w') as f:
    for r in html_data:
      f.write(r+'\n')

  print('Collection exported to {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-export_start_time,2)))


def export_kml(dataframe, filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to KML...')

  kml_data = dataframe[['Location','Long','Lat']]

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

  for _, k in kml_data.iterrows():
    if pd.notnull(k[1]) and pd.notnull(k[2]):
      pnt = kml.newpoint(name=k[0], coords=[(k[1], k[2])], description=k[3])
      pnt.style = style

  kml.save(path=filename)

  print('Collection exported to {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-export_start_time,2)))


def process_holes(holes_database, filename):
  print('Loading {}...'.format(holes_database))
  load_start_time = timeit.default_timer()

  conn = sqlite3.connect(holes_database)
  dataframe = pd.read_sql_query("SELECT * FROM boreholes", conn)

  print('Loaded {} in {} seconds.\n'.format(holes_database, round(timeit.default_timer()-load_start_time,2)))

  dataframe['Long'] = dataframe['Long'].apply(lambda x: None if pd.isnull(x) else round(x, 4))
  dataframe['Lat'] = dataframe['Lat'].apply(lambda x: None if pd.isnull(x) else round(x, 4))
  dataframe['Water_Depth'] = dataframe['Water_Depth'].apply(lambda x: None if pd.isnull(x) else round(x, 2))

  export_csv(dataframe=dataframe, filename=export_filename + '.csv')
  export_html(dataframe=dataframe, filename=export_filename + '.txt')
  export_kml(dataframe=dataframe, filename=export_filename + '.kml')

  conn.close()

  print('Finished processing {} in {} seconds.\n'.format(holes_database, round(timeit.default_timer()-load_start_time,2)))


if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Filter and convert the LacCore Holes Excel file into different formats needed for publishing.')
  parser.add_argument('db_file', type=str, help='Name of CSDCO database file.')
  parser.add_argument('-d', '--no-date', action='store_true', help='Export files without the date in the filename (e.g., collection_YYYYMMDD.csv).')

  args = parser.parse_args()

  if os.path.isfile(args.db_file):
    export_filename = 'collection' if args.no_date else 'collection_' + datetime.datetime.now().strftime('%Y%m%d') 
    process_holes(args.db_file, export_filename)
  else:
    print('ERROR: file \'{}\' does not exist.\n'.format(args.db_file))
