#!/usr/bin/env python
# coding: utf-8

# # Running Dinner Planning 2023
# 
# Code by Levi Soe & Youran

# Importeren van noodzakelijke bibliotheken
import pandas as pd
import random
import xlsxwriter
import os

###  Laad de excel file in met alle bladen (Excelbestand van data)
def load_data(filepath):
    """ Functie om gegevens uit het Excel-bestand te laden """
    # Controleer of het bestand bestaat
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Kan het bestand {filepath} niet vinden.")
    sheets = ["Bewoners", "Adressen", "Paar blijft bij elkaar", "Buren", "Kookte vorig jaar", "Tafelgenoot vorig jaar"]
    return {sheet: pd.read_excel(filepath, sheet, header=0) for sheet in sheets}

### Laad de excel file van de oplosing
def load_oplossing(filepath):
    return pd.read_excel(filepath)

### Pre process bewoners en bladen met incorectte header
def pre_process_dataframes(dataframes):
    """ Voer noodzakelijke voorverwerkingsstappen uit """
    # Bewoners voorverwerken
    dataframes['Bewoners'][['Kookt niet']] = dataframes['Bewoners'][['Kookt niet']].fillna(0).astype(bool)
    # Pas voorverwerking toe voor andere bladen (zoals Paar blijft bij elkaar, Buren, etc.)
    for sheet in ["Paar blijft bij elkaar", "Buren", "Kookte vorig jaar", "Tafelgenoot vorig jaar"]:
        df = dataframes[sheet]
        df.columns = df.iloc[0]
        dataframes[sheet] = df[1:]
    return dataframes

### Evaluatie op basis van penalties
def evaluatie(planning, buren_df, kookte_vorig_jaar_df, tafelgenoot_vorig_jaar_df):
    penalty = 0
    """ Evaluatie code voor improving search """
    # Voorwaarde 1: Maximaal een keer tafelgenoot in de huidige planning
     #In de planning kijken of de bewoners die verschillende namen hebben, 
     # alleen een keer dezelfde value in [voor,hoof,na] kolommen zal hebben.
    
    # Voorwaarde 2: Kookte vorig jaar
    for _, row in kookte_vorig_jaar_df.iterrows():
        huisadres, gang = row['Huisadres'], row['Gang']
        # Als het huis dit jaar nog een hoofd gerecht zal voorbereiden
        if gang == 'Hoofd' and planning.get(huisadres, {}).get('Gang') == gang:
            penalty += 5  # Voeg straf toe
    
    # Voorwaarde 3: Voorkeur gang
    for _, row in adressen_df.iterrows():
        huisadres, voorkeur_gang = row['Huisadres'], row['Voorkeur gang']
        # Als er geen rekening wordt gehouden van de voorkeur gerechten
        if planning.get(huisadres,{}).get('Voorkeur gang') != voorkeur_gang:
            penalty += 5 # Voeg straf toe

    # Voorwaarde 4: Tafelgenoot vorig jaar
    for _, row in tafelgenoot_vorig_jaar_df.iterrows():
        bewoner1, bewoner2 = row['Bewoner1'], row['Bewoner2']
        # Als de bewoners dit jaar weer aan dezelfde tafels zitten
        if planning[bewoner1] == planning[bewoner2]:
            penalty += 5  # Voeg straf toe

    # Voorwaarde 5: Buren
    for _, row in buren_df.iterrows():
        if planning[row['Bewoner1']] == planning[row['Bewoner2']]:  # Als beide buren aan dezelfde tafel zitten
            penalty += 10  # Voeg straf toe
    
    # Voorwaarde 6: 2021 tafelgenoten zijn 
     # we moeten dan weer een dataset laden (dataset 2022 sheet['tefelgenoot vorig jaar])

    
    return penalty


### kijken of de eerste oplossing de eisen voldoen, an=ls het niet voldoen moeten we het veranderen
def  genereer_oplossing(df1, df2):
    """ Genereer een willekeurige planning. """
    planning = {}
    gang = []
    for _, bewoner in df1.iterrows():
        min_groep = df2[df2['Huisadres'] == bewoner['Huisadres']]['Min groepsgrootte'].iloc[0]
        max_groep = df2[ df2['Huisadres']==bewoner['Huisadres']]['Max groepsgrootte'].iloc[0]
        voorkeur_gang = df2[ df2['Huisadres']==bewoner['Huisadres']]['Voorkeur gang'].iloc[0]  #Dit is een wens: zie voorwaarde 3
        if bewoner['Kookt niet']==False:
            huis = random.choice(df2['Huisadres'].tolist())
            if pd.isna(voorkeur_gang):
                gang = random.choice(['Voor', 'Hoofd', 'Na'])
            else:
                gang = voorkeur_gang
            
            planning[bewoner['Bewoner']] = {'voor': huis if gang == 'Voor' else None,
                                            'hoofd': huis if gang == 'Hoofd' else None,
                                            'na': huis if gang == 'Na' else None,
                                            'kookt': gang,
                                            'aantal': None}
        else:
            huis = random.choice(df2['Huisadres'].tolist())
            planning[bewoner['Bewoner']] = {'voor': random.choice(df2['Huisadres'].tolist()), 'hoofd': random.choice(df2['Huisadres'].tolist()), 'na': random.choice(df2['Huisadres'].tolist()), 'kookt': None, 'aantal': None}
    return planning


def planning_eisen(bewoners_df,adressen_df):

    planning = load_oplossing('Running Dinner eerste oplossing 2021.xlsx')

    # Eis 1:Elke deelnemer eet elk gerecht en eet elk gerecht op een ander huisadres.
    for _,row in planning.iterrows():
        voor, hoofd, na = row['Voor'], row['Hoofd'], row['Na']
        # als ze dezelfde huisadres hebben

    # Eis 2:Ieder huishouden dat niet vrijgesteld is van koken, bereidt één van de drie gerechten. 
    for _, row in bewoners_df.iterrows():
         bewoner, huisadres, kookt_niet = row['Bewoner'], row['Huisadres'], row['Kookt niet']
         # Mensen die koken heeft een gerecht
         if kookt_niet==False and planning['kookt'] not in ['Voor','Hoofd','Na']:
            planning['kookt'] == ['Voor','Hoofd','Na']
            gekozen_gerecht = random.choice(['Voor', 'Hoofd', 'Na'])
            planning['kookt'] = gekozen_gerecht

    # Eis 3：Sommige deelnemers hoeven niet te koken en ontvangen op hun huisadres dus voor geen enkele gerecht gasten.
    for _, row in bewoners_df.iterrows():
         bewoner, huisadres, kookt_niet = row['Bewoner'], row['Huisadres'], row['Kookt niet']
         # De mensen die niet koken heeft geen gerecht
         if kookt_niet == True and planning['kookt'] is None:
             print('Eis 3 voldoet') ###3
         else:
            planning['kookt'] = None
        
    # Eis 4: Wanneer een deelnemer een bepaalde gang moet koken is deze deelnemer voor die gang ingedeeld op diens eigen adres. 3#
    def select_column_based_on_kook(row):
        kook_value = row['kookt']
        if kook_value == 'Voor':
            return row['voor']
        elif kook_value == 'Hoofd':
            return row['hoofd']
        elif kook_value == 'Na':
            return row['na']
        else:
            return None  
        
    planning['selected_column'] = planning.apply(select_column_based_on_kook, axis=1)
    

    #Eis 5:	Het aantal tafelgenoten dat op een bepaald huisadres eet, voldoet aan de bij het adres horende minimum en maximum groepsgrootte
    for _, row in bewoners_df.iterrows():
        huisadres = row['Huisadres']
        min_groep = adressen_df[adressen_df['Huisadres'] == bewoner['Huisadres']]['Min groepsgrootte'].iloc[0]
        max_groep = adressen_df[ adressen_df['Huisadres']==bewoner['Huisadres']]['Max groepsgrootte'].iloc[0]
        condition = (planning['aantal'] >= min_groep) & (planning['aantal'] <= max_groep) # Als het de grens overtreed, wordt de adres opnieuw ingepland
    #Eis 6: Een heel klein aantal groepjes van deelnemers, vaak één of twee duo’s, zit tijdens het gehele Running Dinner voor elke gang bij elkaar aan tafel.

    return planning 



### Creeer nieuwe planning
def vind_buur(planning):
    """ Vind een buur door willekeurig een persoon te selecteren en zijn/haar kookgang te veranderen. """
    nieuwe_planning = planning.copy()
    bewoner = random.choice(list(planning.keys()))
    if not planning[bewoner]['kookt']:
        return nieuwe_planning

    nieuwe_gang = random.choice(['Voor', 'Hoofd', 'Na'])
    nieuwe_planning[bewoner]['kookt'] = nieuwe_gang
    return nieuwe_planning


### Improve search 
def hill_climbing(df1, df2, buren, kookte_vorig_jaar, tafelgenoot_vorig_jaar, iteraties):
    """ Zoek naar de beste planning met Hill Climbing. """
    huidige_planning = genereer_oplossing(df1, df2) # is planning eisen
    huidige_score = evaluatie(huidige_planning, buren, kookte_vorig_jaar, tafelgenoot_vorig_jaar)
    
    #evaluatie bij houden
    score_list = []
    score_list.append(huidige_score)

    for _ in range(iteraties):
        nieuwe_planning = vind_buur(huidige_planning)
        nieuwe_score = evaluatie(huidige_planning, buren, kookte_vorig_jaar, tafelgenoot_vorig_jaar)
        
        if nieuwe_score < huidige_score:
            huidige_planning, huidige_score = nieuwe_planning, nieuwe_score
            score_list.append(huidige_score)

    return huidige_planning, score_list


### Sla alle data op in een excel bestand
def save_to_excel(planning, score, df1):
    """ Sla de planning op in een Excel-bestand. """
    rows = []
    for bewoner, details in planning.items():
        huisadres = df1[df1['Bewoner'] == bewoner]['Huisadres'].iloc[0]
        rows.append({
            'Rangnummer': None, # Deze waarde wordt later ingevuld
            'Bewoner': bewoner, 
            'Huisadres': huisadres, 
            'Voor': details['voor'], 
            'Hoofd': details['hoofd'], 
            'Na': details['na'], 
            'kookt': details['kookt'], 
            'aantal': details['aantal']
        })
    
    df_out = pd.DataFrame(rows)
    
    # Voeg oplopende rangnummer toe
    df_out['Rangnummer'] = range(len(df_out))
    
    # Maak een Excel-writer object
    writer = pd.ExcelWriter('Running_Dinner_2023.xlsx', engine='xlsxwriter')

    # Schrijf dataframes naar Excel-bestand
    df_out.to_excel(writer, sheet_name='Planning', index=False)

    #schrijf score 
    score.to_excel(writer, sheet_name='Score', index=False)
    # Sla Excel-bestand op
    writer.close()


### Voor alle functies uit
if __name__ == '__main__':
    # Gebruik de bovenstaande functies
    dfs = load_data('Running Dinner dataset 2021.xlsx')
    df = load_oplossing('Running Dinner eerste oplossing 2021.xlsx')
    preprocessed_dfs = pre_process_dataframes(dfs)
    
    #De verschillende dataframes 
    bewoners_df = preprocessed_dfs['Bewoners']
    adressen_df = preprocessed_dfs['Adressen']
    bij_elkaar_df = preprocessed_dfs['Paar blijft bij elkaar']
    buren_df = preprocessed_dfs['Buren']
    kookte_vorig_jaar_df = preprocessed_dfs['Kookte vorig jaar']
    tafelgenoot_vorig_jaar_df = preprocessed_dfs['Tafelgenoot vorig jaar']
    
    # optimizing search
    beste_planning, score_list = hill_climbing(bewoners_df, 
                                   adressen_df,
                                   buren_df, 
                                   kookte_vorig_jaar_df,
                                   tafelgenoot_vorig_jaar_df,
                                   iteraties=1000)
    
    score_df = pd.DataFrame(score_list)

    save_to_excel(beste_planning, score_df, bewoners_df)

