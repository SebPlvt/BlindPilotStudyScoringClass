# -*- coding: utf-8 -*-
"""
Created on Mon Jul 22 11:35:19 2019

@author: Sébastien Polvent
@contact: sebastien.polvent@unicaen.fr

changelog :
26/07/2019 : added group import
24/07/2016 : added dataframe filter options
23/07/2019 : Initial release, version 1.0
"""

import os
#xlwings v0.15.8
import xlwings as xw
#pandas v0.25.0 with xlrd 1.2.0
import pandas as pd


class BlindPilotStudyScoring():
    """ BlindPilotStudyScoring : computes scores and ratios from an xl_input file
        for "cotation étude pilote_aveugle"
        Save results in a new xl file : cotation étude pilote_aveugle_scorage.xlsx
        (overwrites if file already exists)
    """



    def __init__(self):
        """ init
        sets attributes"""
        os.chdir(os.getcwd())
        print('                     Blind Pilot Study Scoring Script v1.0')
        print('                              Powered by Python3')
        #input file
        self.xl_filename = './cotation étude pilote_aveugle .xlsm'
        #result file
        self.scoring_results_xl_file = './cotation étude pilote_aveugle_scorage.xlsx'
        #group info file
        self.group_filename = './étude pilote_tropes.xlsx'



    def get_file_data(self, filename):
        """get_file_data"""
        print('Reading file info...', end='')
        #open workbook
        xl_file = xw.App(visible=False).books.open(filename)
        i = 1
        nb_sheets = len(xl_file.sheets)
        #get sheets data
        #sheet TOTAL' =  result file template
        subjects_scores = []
        for sh in xl_file.sheets:
            progress = str(round(i/nb_sheets*100))
            print('\rReading file info... ' + progress + '%', end='', flush=True)
            if sh.name == 'TOTAL':
                df_total = pd.read_excel(io=self.xl_filename, sheet_name=sh.name)
            #other sheets = subjects sheets,
            #except 'Modele' but it's not "int" type !
            else:
                try:
                    if isinstance(int(sh.name), int):
                        #retrieve scoring in df
                        tmp_df = pd.read_excel(io=self.xl_filename, \
                                                              sheet_name=sh.name)
                        #avoid empty dataframes
                        #if there is no 'pronoms' column
                        #and if not fully filled required column
                        if not tmp_df.empty and \
                        ", ".join(tmp_df.columns).count('pronoms') and \
                        not tmp_df['type de détail'].isnull().any().any() and \
                        not tmp_df['ordre de la proposition'].isnull().any().any() and \
                        not tmp_df['catégorie de la proposition'].isnull().any().any():
                            subjects_scores.append(tmp_df)
                except ValueError:
                    pass
            i += 1
        #close xl
        xl_file.close()
        for app in xw.apps:
            app.kill()
        #clean columns items names
        df_total.columns = self.clean_list_from_spaces(df_total.columns)
        print('\nData collected !')
        return df_total, subjects_scores



    def clean_list_from_spaces(self, list_to_clean):
        """clean_list_from_spaces, returns a new list"""
        new_list = []
        for item in list_to_clean:
            temp_str = str(item).replace(" ", "_")
            #remove last char if underscore
            if temp_str[len(temp_str)-1:] == "_":
                temp_str = temp_str[:len(temp_str)-1]
            new_list.append(temp_str)
        return new_list



    def data_scoring(self, input_data, df_results):
        """data_scorage, clean data, process scoring, return dataframe of results"""


        def process_scoring(dataframe, serie_index):
            """ process_scoring, returns a pandas serie"""

            result_serie = pd.Series(data=None, index=serie_index)
            result_serie['Sujet'] = int(dataframe.loc[0, 'sujet'])

            #3 memories
            for x in range(1, 4):
                temp_df = dataframe[dataframe['période'] == x]
                #number of propositions
                result_serie["souvenir_" + str(x) + "_nb_propositions"] = \
                    int(temp_df['ordre_de_la_proposition'].describe()['max'])
                #number of internals details
                result_serie["souvenir_" + str(x) + "_nb_détails_internes"] = \
                    len(temp_df[temp_df['type_de_détail'] == "interne"])
                #number of external details
                result_serie["souvenir_" + str(x) + "_nb_détails_externes"] = \
                    len(temp_df[temp_df['type_de_détail'] == "externe"])
                #number of EVE
                result_serie["souvenir_" + str(x) + "_nb_EVE"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'EVE'])
                #number of TPS
                result_serie["souvenir_" + str(x) + "_nb_TPS"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'TPS'])
                #number of L
                result_serie["souvenir_" + str(x) + "_nb_L"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'L'])
                #number of PERC
                result_serie["souvenir_" + str(x) + "_nb_PERC"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'PERC'])
                #number of EMO
                result_serie["souvenir_" + str(x) + "_nb_EMO"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'EMO'])
                #number of SE
                result_serie["souvenir_" + str(x) + "_nb_SE"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'SE'])
                #number of EE
                result_serie["souvenir_" + str(x) + "_nb_EE"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'EE'])
                #number of PS
                result_serie["souvenir_" + str(x) + "_nb_PS"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'PS'])
                #number of GS
                result_serie["souvenir_" + str(x) + "_nb_GS"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'GS'])
                #number of R
                result_serie["souvenir_" + str(x) + "_nb_R"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'R'])
                #number of IO
                result_serie["souvenir_" + str(x) + "_nb_IO"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'IO'])
                #number of FUT
                result_serie["souvenir_" + str(x) + "_nb_FUT"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'FUT'])
                #number of M
                result_serie["souvenir_" + str(x) + "_nb_M"] = \
                    len(temp_df[temp_df['catégorie_de_la_proposition'] == 'M'])

                ##pronouns (for 'type_de_détail' == 'interne' only) :
                intern_df = temp_df[temp_df['type_de_détail'] == "interne"]
                #number of "NA"
                result_serie["souvenir_" + str(x) + "_nb_\"NA\""] = \
                    len(intern_df['pronoms']) - intern_df['pronoms'].count()
                #number of "1"
                result_serie["souvenir_" + str(x) + "_nb_\"1\""] = \
                    len(intern_df[intern_df['pronoms'] == 1])
                #number of "3+6"
                result_serie["souvenir_" + str(x) + "_nb_\"3+6\""] = \
                    len(intern_df[intern_df['pronoms'] == 3]) + \
                    len(intern_df[intern_df['pronoms'] == 6])
                #number of "4"
                result_serie["souvenir_" + str(x) + "_nb_\"4\""] = \
                    len(intern_df[intern_df['pronoms'] == 4])
                #number of "7"
                result_serie["souvenir_" + str(x) + "_nb\"7\""] = \
                len(intern_df[intern_df['pronoms'] == 7])
            #End of process_scoring
            return result_serie


        def process_ratios(dataframe):
            """ process ratios from a dataframe for each memory,
                returns a new dataframe"""

            #3 memories
            for x in range(1, 4):
                #détails_internes
                dataframe["souvenir_" + str(x) + "_ratio_détails_internes"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_internes"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]

                #détails_externes
                dataframe["souvenir_" + str(x) + "_ratio_détails_externes"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]

                ##reference = détails_internes
                #EVE
                dataframe["souvenir_" + str(x) + "_ratio_EVE"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_EVE"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_internes"]
                #EVE_total
                dataframe["souvenir_" + str(x) + "_ratio_EVE_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_EVE"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #TPS
                dataframe["souvenir_" + str(x) + "_ratio_TPS"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_TPS"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_internes"]
                #TPS_total
                dataframe["souvenir_" + str(x) + "_ratio_TPS_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_TPS"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #L
                dataframe["souvenir_" + str(x) + "_ratio_L"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_L"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_internes"]
                #L_total
                dataframe["souvenir_" + str(x) + "_ratio_L_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_L"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #PERC
                dataframe["souvenir_" + str(x) + "_ratio_PERC"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_PERC"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_internes"]
                #PERC_total
                dataframe["souvenir_" + str(x) + "_ratio_PERC_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_PERC"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #EMO
                dataframe["souvenir_" + str(x) + "_ratio_EMO"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_EMO"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_internes"]
                #EMO_total
                dataframe["souvenir_" + str(x) + "_ratio_EMO_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_EMO"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]

                ##reference = détails_externes
                #SE
                dataframe["souvenir_" + str(x) + "_ratio_SE"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_SE"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #SE_total
                dataframe["souvenir_" + str(x) + "_ratio_SE_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_SE"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #EE
                dataframe["souvenir_" + str(x) + "_ratio_EE"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_EE"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #EE_total
                dataframe["souvenir_" + str(x) + "_ratio_EE_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_EE"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #PS
                dataframe["souvenir_" + str(x) + "_ratio_PS"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_PS"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #PS_total
                dataframe["souvenir_" + str(x) + "_ratio_PS_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_PS"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #GS
                dataframe["souvenir_" + str(x) + "_ratio_GS"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_GS"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #GS_total
                dataframe["souvenir_" + str(x) + "_ratio_GS_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_GS"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #R
                dataframe["souvenir_" + str(x) + "_ratio_R"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_R"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #R_total
                dataframe["souvenir_" + str(x) + "_ratio_R_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_R"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #IO
                dataframe["souvenir_" + str(x) + "_ratio_IO"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_IO"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #IO_total
                dataframe["souvenir_" + str(x) + "_ratio_IO_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_IO"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #FUT
                dataframe["souvenir_" + str(x) + "_ratio_FUT"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_FUT"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #FUT_total
                dataframe["souvenir_" + str(x) + "_ratio_FUT_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_FUT"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]
                #M
                dataframe["souvenir_" + str(x) + "_ratio_M"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_M"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_détails_externes"]
                #M_total
                dataframe["souvenir_" + str(x) + "_ratio_M_total"] = \
                    dataframe["souvenir_"  + str(x) + "_nb_M"] / \
                    dataframe["souvenir_"  + str(x) + "_nb_propositions"]

                ##reference = détails_internes
                #interne_"1"
                dataframe["souvenir_" + str(x) + "_ratio_interne_\"1\""] = \
                    dataframe["souvenir_"  + str(x) + "_nb_\"1\""] / \
                    (dataframe["souvenir_"  + str(x) + "_nb_détails_internes"] - \
                     dataframe["souvenir_"  + str(x) + "_nb_\"NA\""])
                #interne_"3+6"
                dataframe["souvenir_" + str(x) + "_ratio_interne_\"3+6\""] = \
                    dataframe["souvenir_"  + str(x) + "_nb_\"3+6\""] / \
                    (dataframe["souvenir_"  + str(x) + "_nb_détails_internes"] - \
                     dataframe["souvenir_"  + str(x) + "_nb_\"NA\""])
                #interne_"4"
                dataframe["souvenir_" + str(x) + "_ratio_interne_\"4\""] = \
                    dataframe["souvenir_"  + str(x) + "_nb_\"4\""] / \
                    (dataframe["souvenir_"  + str(x) + "_nb_détails_internes"] - \
                     dataframe["souvenir_"  + str(x) + "_nb_\"NA\""])
                #interne_"7"
                dataframe["souvenir_" + str(x) + "_ratio_interne_\"7\""] = \
                    dataframe["souvenir_"  + str(x) + "_nb\"7\""] / \
                    (dataframe["souvenir_"  + str(x) + "_nb_détails_internes"] - \
                     dataframe["souvenir_"  + str(x) + "_nb_\"NA\""])
            #End of process_ratios
            return dataframe


        ### main data_scoring code
        print('Scoring data...', end='')
        #clean df_results
        final_df_results = pd.DataFrame(data=None, columns=df_results.columns)
        #clean data
        i = 1
        nb_subjects = len(input_data)
        for subject_df in input_data:
            progress = str(round(i/nb_subjects*100))
            print('\rScoring data... ' + progress + '%', end='', flush=True)
            #clean columns items names
            subject_df.columns = self.clean_list_from_spaces(subject_df.columns)
            #format 'sujet' column -> remove nan values
            subject_df['sujet'] = subject_df.loc[0, 'sujet']
            #format 'période' column -> remove nan values
            clean_tab = []
            for item in subject_df['période']:
                try:
                    if isinstance(int(item), int):
                        epoch = int(item)
                        clean_tab.append(int(item))
                except ValueError:
                    clean_tab.append(epoch)
            subject_df['période'] = clean_tab

            #process scoring
            final_df_results = \
                final_df_results.append(
                    process_scoring(dataframe=subject_df, \
                                    serie_index=final_df_results.columns), \
                                    ignore_index=True)
            #process ratios
            final_df_results = process_ratios(final_df_results)
            i += 1
        #End of data_scoring
        return final_df_results



    def filter_only_data_scores(self, dataframe):
        """filter_only_data_scores, returns filtered dataframe"""
        #dataframe df_scoring filtered with only scores
        df_scores_only = dataframe[[x for x in dataframe.columns \
                                    if str(x).count("_nb_") or \
                                    str(x).count("Sujet")]]
        return  df_scores_only



    def filter_only_data_ratios(self, dataframe):
        """filter_only_data_ratios, returns filtered dataframe"""
        #dataframe df_scoring filtered with only ratios
        df_ratios_only = dataframe[[x for x in dataframe.columns \
                                    if not str(x).count("_nb") or \
                                    str(x).count("Sujet")]]
        return df_ratios_only



    def filter_by_period(self, dataframe, epoch):
        """filter_by_period, epoch must be str even if it's a number,
        returns filtered dataframe"""
        #dataframe df_scoring filtered with only ratios
        df_one_period_only = dataframe[[x for x in dataframe.columns \
                                        if str(x).count("souvenir_" + epoch) or \
                                        str(x).count("Sujet")]]
        return df_one_period_only



    def save_scoring_results(self, df_to_save):
        """save scoring results"""
        print('\nSaving results file...')
        #create and save new xl workbook
        result_file = xw.App(visible=False).books[0]
        result_file.save(path=self.scoring_results_xl_file)
        #copy df_scoring to result file
        result_file.sheets(1).name = 'TOTAL'
        #save xl file
        result_file.save()
        #close xl to avoid conflict between pandas and xlwings
        result_file.close()
        for app in xw.apps:
            app.kill()
        df_to_save.columns = [str(x).replace("\"","'") for x in df_to_save.columns]
        df_to_save.to_excel(excel_writer=self.scoring_results_xl_file, \
                            sheet_name='TOTAL', index=False, startrow=0, startcol=0)
        print('File saved !')



    def open_result_file(self):
        """ open the results file in xl"""
        result_file = xw.Book(self.scoring_results_xl_file)
        #resize columns
        result_file.sheets('TOTAL').autofit()
        result_file.save()



    def add_group_info_to_df(self, dataframe, group_filename):
        """ add_group_info in the result dataframe
        to make statistics and visualizations
        returns the dataframe with the group info"""
        #read group file info
        groups_info = pd.read_excel(io=group_filename, sheet_name='pourcentages')
        #drop duplicates
        groups_info = groups_info.drop_duplicates(subset='sujet ')
        #drop if NA in columns ['groupe', 'sujet ', 'code étude', 'code Remember']
        groups_info = groups_info.dropna(axis=0, subset=['groupe', 'sujet ', \
                                                         'code étude', 'code Remember'])
        #put subject in index
        groups_info.index = [int(x) for x in groups_info['sujet ']]
        dataframe.index = [int(x) for x in dataframe['Sujet']]

        #retrieve group info for each subject
        dataframe.insert(loc=1, column='Groupe', value=None)
        dataframe.insert(loc=1, column='code_étude', value=None)
        dataframe.insert(loc=1, column='code_Remember', value=None)
        for sujet1 in dataframe.index:
            for sujet2 in groups_info.index:
                if sujet1 == sujet2:
                    dataframe.at[sujet1, 'Groupe'] = \
                        groups_info.at[sujet1, 'groupe']
                    dataframe.at[sujet1, 'code_étude'] = \
                        groups_info.at[sujet1, 'code étude']
                    dataframe.at[sujet1, 'code_Remember'] = \
                        groups_info.at[sujet1, 'code Remember']

        return dataframe
