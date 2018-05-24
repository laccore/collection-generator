import pandas as pd
import datetime
import simplekml
import timeit
import argparse
import os.path

pd.options.mode.chained_assignment = None

def export_csv(dataframe, export_filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to CSV...')
  csv_data = dataframe.loc[:,['Location','Country','State_Province','Hole ID','Original ID','Date','Water Depth (m)','Lat','Long','Elevation','Position','Sample Type','mblf T','mblf B']]
  csv_data.Date = csv_data.Date.apply(lambda x: x.strftime('%Y%m%d') if isinstance(x, pd.datetime) and pd.notnull(x) else x)

  csv_data.to_csv(export_filename, encoding='utf-8-sig', index=False)

  print('Collection exported to {} in {} seconds.\n'.format(export_filename, round(timeit.default_timer()-export_start_time,2)))


def export_html(dataframe, export_filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to HTML...')
  html_data = '<tr><td>' + dataframe.loc[:,'Location'] + '</td><td>'
  html_data += dataframe.loc[:,'Lat'].map(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe.loc[:,'Long'].map(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe.loc[:,'Elevation'].map(lambda x: '' if pd.isnull(x) else str(round(x))) + '</td></tr>'

  with open(export_filename, 'w') as f:
    for r in html_data:
      f.write(r+'\n')

  print('Collection exported to {} in {} seconds.\n'.format(export_filename, round(timeit.default_timer()-export_start_time,2)))


def export_kml(dataframe, export_filename):
  export_start_time = timeit.default_timer()
  print('Exporting collection to KML...')

  kml_data = dataframe.loc[:,['Location','Long','Lat','Google Earth Comment Field']]
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


def process_holes(holes_filename, export_filename):
  print('Loading {}...'.format(holes_filename))
  load_start_time = timeit.default_timer()
  df = pd.read_excel(holes_filename, parse_dates=False)

  df.loc[:, ['Lat', 'Long']] = df.loc[:,['Lat', 'Long']].apply(lambda x: round(x, 4))
  df.loc[:, 'Water Depth (m)'] = df.loc[:,'Water Depth (m)'].apply(lambda x: round(x, 2))

  print('Loaded {} in {} seconds.\n'.format(holes_filename, round(timeit.default_timer()-load_start_time,2)))

  export_csv(df, export_filename + '.csv')
  export_html(df, export_filename + '.txt')
  export_kml(df, export_filename + '.kml')

  print('Finished processing {} in {} seconds.\n'.format(holes_filename, round(timeit.default_timer()-load_start_time,2)))


if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Filter and convert the LacCore Holes Excel file into different formats needed for publishing.')
  parser.add_argument('holes_file', type=str, help='Name of LacCore Holes file.')
  parser.add_argument('-d', action='store_true', help='Export files with the date in the filename (format: YYYYMMDD).')

  args = parser.parse_args()

  if os.path.isfile(args.holes_file):
    filename = 'collection_' + datetime.datetime.now().strftime('%Y%m%d') if args.d else 'collection'
    process_holes(args.holes_file, filename)
  else:
    print("ERROR: file '{}' does not exist or is not readable.\n".format(args.holes_file))
