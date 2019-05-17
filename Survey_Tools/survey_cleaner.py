"""
Reads a csv containing a stream survey, parses it, then writes
the parsed data to xlsx files.

Works with pyfluv v0.3.3.post2, pandas 0.23.4, Python 3.7

Script written 5/16/2019
"""
import os

from pandas import ExcelWriter
import pyfluv

##### Set these variables #####



file_location = r'C:\Users\sky.jones\Desktop\nstTest'
file_name = r'wpr_myr5_adjusted' # must be a .csv file, but don't write the extension.

xs_append = "_CROSSES"
pro_append = "_PROFILES"



##### End setup #####

file = os.path.join(file_location, file_name+'.csv')

survey = pyfluv.StreamSurvey(file=file, sep=',', metric=False, keywords=None,
                             colRelations=None)

crosses = survey.get_cross_objects(guessType=True, project=True, stripName=False)
pros = survey.get_profile_objects(stripName=False)

cross_file = os.path.join(file_location, file_name+xs_append+'.xlsx')
pro_file = os.path.join(file_location, file_name+pro_append+'.xlsx')

cross_dfs = {cross.name:cross.df for cross in crosses}

wanted_cols = ['shotnum','exes','whys','zees','desc','Thalweg','Station',
               'Riffle','Run', 'Pool','Glide','Water Surface','Bankfull',
               'Top of Bank','Cross Section','Structure']
pro_dfs = {}
for pro in pros:
    have_cols = [l for l in wanted_cols if l in pro.filldf.columns]
    pro_dfs[pro.name] = pro.filldf[have_cols]

writer = ExcelWriter(cross_file)
for key, df in cross_dfs.items():
    df.to_excel(writer, '%s' % key)
writer.save()

writer = ExcelWriter(pro_file)
for key, df in pro_dfs.items():
    df.to_excel(writer, '%s' % key)
writer.save()

