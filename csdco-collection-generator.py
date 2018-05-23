import pandas as pd
import datetime
import simplekml
import timeit
import argparse
import os.path

pd.options.mode.chained_assignment = None

def export_csv(dataframe):
  export_start_time = timeit.default_timer()
  print('Exporting collection to CSV...')
  csv_data = dataframe.loc[:,['Location','Country','State_Province','Hole ID','Original ID','Date','Water Depth (m)','Lat','Long','Elevation','Position','Sample Type','mblf T','mblf B']]
  csv_data.Date = csv_data.Date.apply(lambda x: x.strftime('%Y%m%d') if isinstance(x, pd.datetime) else x)

  filename = 'collection_' + datetime.datetime.now().strftime('%Y%m%d') + '.csv'
  csv_data.to_csv(filename, encoding='utf-8', index=False)

  print('Collection exported to {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-export_start_time,2)))


def export_html(dataframe):
  export_start_time = timeit.default_timer()
  print('Exporting collection to HTML...')
  html_data = '<tr><td>' + dataframe.loc[:,'Location'] + '</td><td>'
  html_data += dataframe.loc[:,'Lat'].map(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe.loc[:,'Long'].map(lambda x: '' if pd.isnull(x) else str(x)) + '</td><td>'
  html_data += dataframe.loc[:,'Elevation'].map(lambda x: '' if pd.isnull(x) else str(round(x))) + '</td></tr>'

  filename = 'collection_' + datetime.datetime.now().strftime('%Y%m%d') + '.txt'
  with open(filename, 'w') as f:
    for r in html_data:
      f.write(r+'\n')

  print('Collection exported to {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-export_start_time,2)))


def export_kml(dataframe):
  export_start_time = timeit.default_timer()
  print('Exporting collection to KML...')

  kml_data = dataframe.loc[:,['Location','Long','Lat','Google Earth Comment Field']]
  kml = simplekml.Kml(name='LacCore/CSDCO Core Collection')
  style = simplekml.Style()
  style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png'
  style.iconstyle.scale = 1.2
  style.iconstyle.color = 'ff0000cc'

  for i, k in kml_data.iterrows():
    pnt = kml.newpoint(name=k[0], coords=[(k[1], k[2])], description=k[3])
    pnt.style = style
  filename = 'collection_' + datetime.datetime.now().strftime('%Y%m%d') + '.kml'

  kml.save(path=filename)

  print('Collection exported to {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-export_start_time,2)))


def process_holes(filename):
  print('Loading {}...'.format(filename))
  load_start_time = timeit.default_timer()
  df = pd.read_excel(filename, parse_dates=False)

  df.loc[:, ['Lat', 'Long']] = df.loc[:,['Lat', 'Long']].apply(lambda x: round(x, 4))
  df.loc[:, 'Water Depth (m)'] = df.loc[:,'Water Depth (m)'].apply(lambda x: round(x, 2))

  print('Loaded {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-load_start_time,2)))

  export_csv(df)
  export_html(df)
  export_kml(df)

  print('Finished processing {} in {} seconds.\n'.format(filename, round(timeit.default_timer()-load_start_time,2)))


if __name__ == '__main__':
  parser = argparse.ArgumentParser(description='Filter and convert the LacCore Holes Excel file into different formats needed for publishing.')
  parser.add_argument('holes_file', type=str, help='Name of LacCore Holes file.')

  args = parser.parse_args()

  if os.path.isfile(args.holes_file):
    process_holes(args.holes_file)
  else:
    print("ERROR: file '{}' does not exist or is not readable.\n".format(args.holes_file))
