#openpyxl pandas
#223.宝可梦技能强化表@PokemonSkillFortify 26.宝可梦天赋方案表@Talent_Plan 25.宝可梦天赋表@Pokemon_Talent

import pandas as pd
import math
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

RESULT_GROUP_ID = 10000
RESULTDF = pd.DataFrame(columns=['ID', 'UnityID', 'UnityName', 'TriggerLevel', 'Skill1', 'Skill2', 'Talen1', 'Talen2', 'UTanlen'])
RESULTDF_INDEX = 0

def write_to_sheet(df, excel_file_path):
    with pd.ExcelWriter(excel_file_path) as writer:
        df.to_excel(writer, sheet_name='result')


def read_excel_file(file_name,keep_list, check_key):
    df = pd.read_excel(file_name)
    df = df[keep_list]
    df = df[df[check_key].isin(SkillSet['ID'])]
    return df

# def ApendRow(UnityID, GroupID, TriggerLevel, SkillType, SkillID):
#     global RESULTDF_INDEX
#     RESULTDF.loc[RESULTDF_INDEX] = [UnityID, GroupID, TriggerLevel, SkillType, SkillID]
#     RESULTDF_INDEX += 1

def ApendRow(UnityID, UnityName, TriggerLevel, Skill1,  Skill2, Talen1, Talen2, UTanlen):
    global RESULTDF_INDEX
    RESULTDF.loc[RESULTDF_INDEX] = [RESULTDF_INDEX+1, UnityID, UnityName, TriggerLevel, Skill1, Skill2, Talen1, Talen2, UTanlen]
    RESULTDF_INDEX += 1

def gen_hero_data(row):
    global RESULT_GROUP_ID
    RESULT_GROUP_ID += 1
    # skill1_slot = 0
    # skill1_branch = math.floor((row.skill1-1) / 2 + 1)

    # #获取1技能
    # filtered_df = PokemonTallent[(PokemonTallent['GroupId']==row.ID) & (PokemonTallent['SkillIndexInSlot']==skill1_slot) & (PokemonTallent['BranchIndex']==skill1_branch)]
    # result = filtered_df.sort_values('TriggerLevel').iloc[0]
    # ApendRow(row.ID, RESULT_GROUP_ID, result['TriggerLevel'], '1', result['TalentId'])

    # #获取2技能
    # skill2_slot = 1
    # skill2_branch = math.floor((row.skill2-1) / 2 + 1)

    # filtered_df = PokemonTallent[(PokemonTallent['GroupId']==row.ID) & (PokemonTallent['SkillIndexInSlot']==skill2_slot) & (PokemonTallent['BranchIndex']==skill2_branch)]
    # result = filtered_df.sort_values('TriggerLevel', ascending=False).iloc[0]    
    # ApendRow(row.ID, RESULT_GROUP_ID, result['TriggerLevel'], '1', result['TalentId'])

    # skill3_slot = 3
    # skill3_branch = 1

    # #获取集结技能
    # filtered_df = PokemonTallent[(PokemonTallent['GroupId']==row.ID) & (PokemonTallent['SkillIndexInSlot']==skill3_slot) & (PokemonTallent['BranchIndex']==skill3_branch)]
    # result = filtered_df.sort_values('TriggerLevel').iloc[0]
    # ApendRow(row.ID, RESULT_GROUP_ID, result['TriggerLevel'], '1', result['TalentId'])

    # fortify_slot_index = 0
    # fortify_slot = math.floor((row.skill1-1) / 2 + 1)
    # fortify_branch = 2 - (row.skill1 % 2)
    # filtered_df = PokemonSkillFortify[(PokemonSkillFortify['UnitId']==row.ID) & (PokemonSkillFortify['SkillSlotIndex']==fortify_slot_index) & (PokemonSkillFortify['SkillBranch']==fortify_slot) 
    #                                   & (PokemonSkillFortify['SkillFortifyIndex']==fortify_branch)]
    # result = filtered_df.sort_values('NextLevelUnlockLimit').iloc[0]
    # ApendRow(row.ID, RESULT_GROUP_ID, result['NextLevelUnlockLimit'], '2', result['Id'])

    # fortify_slot_index = 1
    # fortify_slot = math.floor((row.skill2-1) / 2 + 1)
    # fortify_branch = 2 - (row.skill2 % 2)
    # filtered_df = PokemonSkillFortify[(PokemonSkillFortify['UnitId']==row.ID) & (PokemonSkillFortify['SkillSlotIndex']==fortify_slot_index) & (PokemonSkillFortify['SkillBranch']==fortify_slot) 
    #                                   & (PokemonSkillFortify['SkillFortifyIndex']==fortify_branch)]
    # result = filtered_df.sort_values('NextLevelUnlockLimit').iloc[0]
    # ApendRow(row.ID, RESULT_GROUP_ID, result['NextLevelUnlockLimit']+2, '2', result['Id'])

    # fortify_slot_index = 3
    # fortify_slot = row.skill3
    # fortify_branch = row.skill3
    # filtered_df = PokemonSkillFortify[(PokemonSkillFortify['UnitId']==row.ID) & (PokemonSkillFortify['SkillSlotIndex']==fortify_slot_index) & (PokemonSkillFortify['SkillBranch']==fortify_slot) 
    #                                   & (PokemonSkillFortify['SkillFortifyIndex']==fortify_branch)]
    # result = filtered_df.sort_values('NextLevelUnlockLimit').iloc[0]
    # ApendRow(row.ID, RESULT_GROUP_ID, result['NextLevelUnlockLimit'], '2', result['Id'])
    ApendRow(row.ID, row.Name, 1, 1, 1, 1, 1, 1)

    level = 1
    p = 1
    while(p <= 2):
        filtered_df = SkillRecommend[(SkillRecommend['Priority']== p ) & (SkillRecommend['PokemonId'] ==row.HeroID)]

        print(row.ID)
        if(filtered_df.empty == False):
            OneFortify_df = PokemonSkillFortify[(PokemonSkillFortify['Id'] == filtered_df['OneFortifyID'].iloc[0])]
            TwoFortifyID_df = PokemonSkillFortify[(PokemonSkillFortify['Id'] == filtered_df['TwoFortifyID'].iloc[0])]
            UFortifyID_df = PokemonSkillFortify[(PokemonSkillFortify['Id'] == filtered_df['UFortifyID'].iloc[0])]

            if(OneFortify_df['SkillBranch'].iloc[0] != 1 | OneFortify_df['SkillFortifyIndex'].iloc[0] != 1):
                level += 1
            if(TwoFortifyID_df['SkillBranch'].iloc[0] != 1 | TwoFortifyID_df['SkillFortifyIndex'].iloc[0] != 1):
                level += 1
            if(UFortifyID_df['SkillFortifyIndex'].iloc[0] != 1):
                level += 1

        if(level > 1):
            ApendRow(row.ID, row.Name,level, OneFortify_df['SkillBranch'].iloc[0],  OneFortify_df['SkillFortifyIndex'].iloc[0], TwoFortifyID_df['SkillBranch'].iloc[0],  TwoFortifyID_df['SkillFortifyIndex'].iloc[0], UFortifyID_df['SkillFortifyIndex'].iloc[0])
            break
        
        p += 1


    # ApendRow(row.ID, row.Name,level, OneFortify_df['SkillBranch'].iloc[0],  OneFortify_df['SkillFortifyIndex'].iloc[0], TwoFortifyID_df['SkillBranch'].iloc[0],  TwoFortifyID_df['SkillFortifyIndex'].iloc[0], UFortifyID_df['SkillFortifyIndex'].iloc[0])


if __name__ == "__main__":
    SkillSet = pd.read_excel('PokemonSkillSet.xlsx')

    # PokemonHero = read_excel_file('11.宝可梦实体配置@Pokemon_Hero.xlsx', ['UnitId', 'ToolRemark1'],'UnitId')
    PokemonTallent = read_excel_file('26.宝可梦天赋方案表@Talent_Plan.xlsx', ['Id', 'GroupId', 'TriggerLevel', 'ChooseType', 'TalentId', 'SkillIndexInSlot', 'BranchIndex'],'GroupId')
    PokemonSkillFortify = read_excel_file('223.宝可梦技能强化表@PokemonSkillFortify.xlsx', ['Id', 'UnitId', 'SkillSlotType', 'SkillSlotIndex', 'SkillBranch', 'SkillFortifyIndex', 'NextLevelUnlockLimit'],'UnitId')
    SkillRecommend = pd.read_excel('宝可梦招式套路.xlsx')
    SkillRecommend = SkillRecommend[['PokemonId', 'Priority', 'OneFortifyID', 'TwoFortifyID', 'UFortifyID']]


    for row in SkillSet.itertuples():
        gen_hero_data(row)

    print(RESULTDF)
    write_to_sheet(RESULTDF, 'Result.xlsx')