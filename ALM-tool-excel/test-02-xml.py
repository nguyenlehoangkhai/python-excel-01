import pandas as pd

def locate(data, query, value, output):
    df = pd.DataFrame(data = data)
    # create a list of values in the query (column)
    values = df[query].tolist()
    print('values'+ '\n', values, '\n')
    row = 0
    if value in values:
        row = values.index(value)
    return df.loc[row, output]
d = {
    'AtomicNumber': [1, 2, 3, 4, 5],
    'Element': ['Hydrogen', 'Helium', 'Lithium', 'Beryllium', 'Boron'],
    'Symbol': ['H', 'He', 'Li', 'Be', 'B'],
    'BoilingPoint': [20.28, 4.22, 1615.00, 2742.00, 4200.00],
}
#
value = locate(data=d, query='Symbol', value='He', output='BoilingPoint')

print(type(d), '\n')

print('Dictionary' + '\n', d, '\n')



df = pd.DataFrame(d)
print('Dataframe' + '\n', df)


'''df.to_excel('data.xlsx', index=False)
df.to_json('data.js')

def func(row):
    xml = ['<item>']
    for field in row.index:
        xml.append('  <{0}>{1}</{0}>'.format(field, row[field]))
    xml.append('</item>')
    return '\n'.join(xml)

print('\n'.join(df.apply(func, axis=1)))
'''
# df.to_excel('xml_data.xml')
print(df.query('Element == "Helium"')['BoilingPoint'])
