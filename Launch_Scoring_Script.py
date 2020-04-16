# -*- coding: utf-8 -*-
"""
Created on Mon Jul 22 11:35:19 2019

@author: Sébastien Polvent
@contact: sebastien.polvent@unicaen.fr

changelog :
24/07/2016 : added dataframe filter options
23/07/2019 : Initial release, version 1.0
"""

import os
import time
import U1077LogoClass
import BlindPilotStudyScoringClass as BPSSC

U1077LogoClass.U1077_Logo().print_logo()
#time marker, start of processing
TIMER = time.time()


### CHOOSE TO SHOW GROUP INFO IN RESULT FILE (True to show, False to hide) ###
SHOW_GROUP_INFO = True
### CHOOSE TO SHOW GROUP INFO IN RESULT FILE       /!\ ###


#create object
study_scoring = BPSSC.BlindPilotStudyScoring()
print("\nPlease wait while processing...", flush=True)

#get file data
df_scoring, subject_scores = \
    study_scoring.get_file_data(filename=study_scoring.xl_filename)
#process data scoring
df_scoring = study_scoring.data_scoring(input_data=subject_scores, \
                                        df_results=df_scoring)
#add group info : send copy to avoid modification of original dataframe
df_scoring_groups = study_scoring.add_group_info_to_df(dataframe=df_scoring[:], \
                                                       group_filename=study_scoring.group_filename)


#dataframe df_scoring_groups filtered with only scores
df_scoring_scores_only = study_scoring.filter_only_data_scores(dataframe=df_scoring_groups)
#dataframe df_scoring_groups filtered with only ratios
df_scoring_ratios_only = study_scoring.filter_only_data_ratios(dataframe=df_scoring_groups)


#dataframes filtered by epoch
df_souvenir1 = study_scoring.filter_by_period(dataframe=df_scoring_groups, \
                                              epoch="1")
df_souvenir2 = study_scoring.filter_by_period(dataframe=df_scoring_groups, \
                                              epoch="2")
df_souvenir3 = study_scoring.filter_by_period(dataframe=df_scoring_groups, \
                                              epoch="3")
#then dataframes filtered by scores and ratios for each epoch
#souvenir 1
df_souvenir1_scores = study_scoring.filter_only_data_scores(dataframe=df_souvenir1)
df_souvenir1_ratios = study_scoring.filter_only_data_ratios(dataframe=df_souvenir1)
#souvenir 2
df_souvenir2_scores = study_scoring.filter_only_data_scores(dataframe=df_souvenir2)
df_souvenir2_ratios = study_scoring.filter_only_data_ratios(dataframe=df_souvenir2)
#souvenir 3
df_souvenir3_scores = study_scoring.filter_only_data_scores(dataframe=df_souvenir3)
df_souvenir3_ratios = study_scoring.filter_only_data_ratios(dataframe=df_souvenir3)
#you can also first filtering by epoch and then by scores_only or ratio_only !


#save in a new xl file
if SHOW_GROUP_INFO:
    study_scoring.save_scoring_results(df_to_save=df_scoring_groups)
else:
    study_scoring.save_scoring_results(df_to_save=df_scoring)
#open the newly created file
study_scoring.open_result_file()
#compute time to process the script
TEMPS_INTER = time.time() - TIMER
print('\rAnalysis Completed !')
print(f'Processing time : {TEMPS_INTER:.2f} seconds.')
os.system("pause")
