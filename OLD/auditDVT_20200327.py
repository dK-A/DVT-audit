import pandas as pd
import tkinter as tk

# folder_path
folder_path = './'

# import csv files
inpatients_1 = pd.read_csv(folder_path + 'data/0患者情報/病床管理_入院・転入一覧.csv', skiprows=[0, 1, 2, 3, 4], parse_dates=[4, 26],
                           encoding='cp932')
inpatients_2 = pd.read_csv(folder_path + 'data/0患者情報/病床管理_転科・担当変更一覧.csv', skiprows=[0, 1, 2, 3], parse_dates=[3, 12],
                           encoding='cp932')
template_op = pd.read_csv(folder_path + 'data/7DVT/静脈血栓塞栓症リスク評価.csv', skiprows=[0, 1], parse_dates=[0, 1, 2, 22, 23],
                          encoding='cp932')
template_others = pd.read_csv(folder_path + 'data/7DVT/非手術例の静脈血栓塞栓症リスク評価.csv', skiprows=[0, 1],
                              parse_dates=[0, 1, 2, 22, 23], encoding='cp932')


# periods check
# 病床管理_入院・転入一覧の期間
admission_periods = pd.read_csv(folder_path + 'data/0患者情報/病床管理_入院・転入一覧.csv', nrows=2, encoding='cp932')
admission_date = admission_periods.iloc[0, 2]

# 病床管理_転科・担当変更一覧の期間
transferred_periods = pd.read_csv(folder_path + 'data/0患者情報/病床管理_転科・担当変更一覧.csv', nrows=2, encoding='cp932')
transferred_date = transferred_periods.iloc[0, 2]


def start_end_date(periods):
    start_date = periods.split('(')[0]
    end_date = periods.split('(')[1]
    end_date = end_date.split(' ')[2]
    periods = [start_date, end_date]
    return periods


admission_date = start_end_date(admission_date)
audit_admission_periods = ['入院・転入患者'] + admission_date
transferred_date = start_end_date(transferred_date)
audit_transferred_periods = ['転科・担当変更患者'] + transferred_date


def start_end_date2(periods):
    start_date = periods.split('～')[0]
    end_date = periods.split('～')[1]
    periods = [start_date, end_date]
    return periods


# 静脈血栓塞栓症リスク評価の抽出期間
template_op_periods = pd.read_csv(folder_path + 'data/7DVT/非手術例の静脈血栓塞栓症リスク評価.csv', nrows=1, encoding='cp932')
template_op_date = template_op_periods.columns[0]
template_op_date = start_end_date2(template_op_date)
audit_template_op_periods = ['静脈血栓塞栓症リスク評価'] + template_op_date

# 非手術例の静脈血栓塞栓症リスク評価の抽出期間
template_others_periods = pd.read_csv(folder_path + 'data/7DVT/非手術例の静脈血栓塞栓症リスク評価.csv', nrows=1, encoding='cp932')
template_others_date = template_others_periods.columns[0]
template_others_date = start_end_date2(template_others_date)
audit_template_others_periods = ['非手術例の静脈血栓塞栓症リスク評価'] + template_others_date

# 集計
audit_periods = [audit_admission_periods, audit_transferred_periods, audit_template_op_periods,
                 audit_template_others_periods]
audit_periods = pd.DataFrame(audit_periods, columns=['監査', '開始日', '終了日'])


# 西暦変換
def seireki_change(date):
    if date[0] == 'H':
        Y = int(date[1:3]) + 1988
        MD = date[4:]
        seireki_date = str(Y) + '/' + MD
        seireki_date = seireki_date.replace(' ', '')
    elif date[0] == 'R':
        if date[1] == ' ':
            Y = int(date[2]) + 2018
            MD = date[4:]
            seireki_date = str(Y) + '/' + MD
            seireki_date = seireki_date.replace(' ', '')
        else:
            Y = int(date[1:3]) + 2018
            MD = date[4:]
            seireki_date = str(Y) + '/' + MD
            seireki_date = seireki_date.replace(' ', '')
    else:
        seireki_date = date.replace(' ', '')
    return seireki_date


for i in range(len(audit_periods)):
    s_date = audit_periods.loc[i, '開始日']
    e_date = audit_periods.loc[i, '終了日']
    start_date = seireki_change(s_date)
    end_date = seireki_change(e_date)
    audit_periods.loc[i, '開始日'] = start_date
    audit_periods.loc[i, '終了日'] = end_date

audit_periods.to_csv(folder_path + 'data/10監査期間/患者データ・テンプレートの抽出期間.csv', encoding='utf-8_sig')


# DataFrameの成型 inpatients, template_op, template_others
inpatients_1 = inpatients_1.drop(
    inpatients_1.columns[[0, 2, 3, 5, 6, 7, 8, 13, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27]], axis=1)
inpatients_1 = inpatients_1[inpatients_1['登録区分'] == '入院確定']
inpatients_1['日時'] = inpatients_1['日時'].str.split('(', expand=True)[0]
inpatients_1 = inpatients_1.rename(columns={'日時': '日付', '患者コード': '患者ID'})
inpatients_2 = inpatients_2.drop(inpatients_2.columns[[0, 2, 6, 11, 12, 13]], axis=1)
inpatients_2['日時'] = inpatients_2['日時'].str.split('(', expand=True)[0]
inpatients_2['科・医師'] = inpatients_2['科・医師'].str.split(' ⇒ ', expand=True)[1]
inpatients_2['入院診療科'] = inpatients_2['科・医師'].str.split('・', expand=True)[0]
inpatients_2['入院主治医'] = inpatients_2['科・医師'].str.split('・', expand=True)[1]
inpatients_2 = inpatients_2.rename(columns={'日時': '日付', '患者コード': '患者ID'})
inpatients = pd.concat([inpatients_1, inpatients_2], sort=False, ignore_index=True)
inpatients = inpatients[['登録区分', '日付', '患者ID', '氏名', '性別', '入院診療科', '入院主治医']]
inpatients.to_csv(folder_path + 'data/13監査結果/入院患者DVT.csv', encoding='utf-8_sig')

# 手術例のDVTリスク評価
template_op = template_op.iloc[:,
              [0, 6, 7, 8, 21, 35, 38, 41, 44, 47, 50, 53, 56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, 89, \
               92, 95, 98, 101, 104, 107, 110, 113, 116, 119, 122, 125, 128, 131, 134, 137, 140, 143, 146, 149, \
               152, 155, 158, 161, 164, 167, 170, 176, 179, 182, 185, 188, 191, 194, 197, 251]]
template_op = template_op.rename(
    columns={'Unnamed: 140': '不加的リスク', '診療科': '入院診療科', '可能な範囲で早期離床や積極的な運動を行う': '早期離床', 'ベッド上での運動療法（リハビリテーションを依頼）': 'リハビリ', 'Unnamed: 251': 'その他'})

template_op.loc[template_op['45分以内の手術（低リスク）'] == 1, 'リスク評価1'] = 1
template_op.loc[template_op['45分を超える手術（中リスク）'] == 1, 'リスク評価1'] = 2
template_op.loc[template_op['45分以内の検査・処置（低リスク）'] == 1, 'リスク評価1'] = 1
template_op.loc[template_op['45分を超える検査・処置（中リスク）'] == 1, 'リスク評価1'] = 2
template_op.loc[template_op['上肢手術（低リスク）'] == 1, 'リスク評価2'] = 1
template_op.loc[template_op['人工股関節置換術（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['脊椎手術（中リスク）'] == 1, 'リスク評価2'] = 2
template_op.loc[template_op['人工膝関節置換術（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['下肢手術（中リスク）'] == 1, 'リスク評価2'] = 2
template_op.loc[template_op['股関節骨折手術（高リスク）。　（大腿骨骨幹部を含む）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['上肢手術（中リスク）。（腸骨からの採骨、下肢から。　神経・皮膚の採取を伴う）'] == 1, 'リスク評価2'] = 2
template_op.loc[template_op['下肢悪性腫瘍手術（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['骨盤骨切り術（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['骨盤骨折（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['脊椎・脊髄損傷（中リスク）'] == 1, 'リスク評価2'] = 2
template_op.loc[template_op['重症外傷（多発外傷）（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['大腿骨遠位部以下の。単独外傷（中リスク）'] == 1, 'リスク評価2'] = 2
template_op.loc[template_op['重症熱傷／重症脊損（高リスク）'] == 1, 'リスク評価2'] = 3
template_op.loc[template_op['非開頭術（低リスク）'] == 1, 'リスク評価3'] = 1
template_op.loc[template_op['開頭術（中リスク）'] == 1, 'リスク評価3'] = 2
template_op['リスク評価1'] = template_op['リスク評価1'].fillna(0)
template_op['リスク評価2'] = template_op['リスク評価2'].fillna(0)
template_op['リスク評価3'] = template_op['リスク評価3'].fillna(0)

template_op['不加的リスク'] = template_op['不加的リスク'].fillna('@')
template_op.loc[template_op['不加的リスク'] == '@', 'リスク評価不加'] = 200
template_op.loc[template_op['不加的リスク'] == -2, 'リスク評価不加'] = -1
template_op.loc[template_op['不加的リスク'] == -1, 'リスク評価不加'] = 0
template_op.loc[template_op['不加的リスク'] == 0, 'リスク評価不加'] = 0
template_op.loc[template_op['不加的リスク'] == 1, 'リスク評価不加'] = 0
template_op.loc[template_op['不加的リスク'] == 2, 'リスク評価不加'] = 1
template_op.loc[template_op['不加的リスク'] == 3, 'リスク評価不加'] = 1
template_op.loc[template_op['不加的リスク'] == 4, 'リスク評価不加'] = 2
template_op.loc[template_op['不加的リスク'] == 5, 'リスク評価不加'] = 2
template_op.loc[template_op['不加的リスク'] == 6, 'リスク評価不加'] = 2
template_op.loc[template_op['不加的リスク'] == 7, 'リスク評価不加'] = 3
template_op['リスク評価不加'] = template_op['リスク評価不加'].fillna(3)

template_op['リスク評価0'] = template_op[['リスク評価1', 'リスク評価2', 'リスク評価3']].max(axis=1)
template_op['リスク合計点'] = template_op['リスク評価0'] + template_op['リスク評価不加']
template_op.loc[template_op['リスク合計点'] <= 0, 'リスク評価'] = 'リスクなし'
template_op.loc[template_op['リスク合計点'] == 1, 'リスク評価'] = '低リスク'
template_op.loc[template_op['リスク合計点'] == 2, 'リスク評価'] = '中リスク'
template_op.loc[template_op['リスク合計点'] == 3, 'リスク評価'] = '高リスク'
template_op.loc[template_op['リスク合計点'] == 4, 'リスク評価'] = '最高リスク'
template_op.loc[template_op['リスク合計点'] == 5, 'リスク評価'] = '最高リスク'
template_op.loc[template_op['リスク合計点'] == 6, 'リスク評価'] = '最高リスク'

template_op.loc[template_op['45分以内の手術（低リスク）'] == 1, '45分以内の手術（低リスク）'] = '45分以内の手術（低リスク）'
template_op.loc[template_op['45分を超える手術（中リスク）'] == 1, '45分を超える手術（中リスク）'] = '45分を超える手術（中リスク）'
template_op.loc[template_op['45分以内の検査・処置（低リスク）'] == 1, '45分以内の検査・処置（低リスク）'] = '45分以内の検査・処置（低リスク）'
template_op.loc[template_op['45分を超える検査・処置（中リスク）'] == 1, '45分を超える検査・処置（中リスク）'] = '45分を超える検査・処置（中リスク）'
template_op.loc[template_op['上肢手術（低リスク）'] == 1, '上肢手術（低リスク）'] = '上肢手術（低リスク）'
template_op.loc[template_op['人工股関節置換術（高リスク）'] == 1, '人工股関節置換術（高リスク）'] = '人工股関節置換術（高リスク）'
template_op.loc[template_op['脊椎手術（中リスク）'] == 1, '脊椎手術（中リスク）'] = '脊椎手術（中リスク）'
template_op.loc[template_op['人工膝関節置換術（高リスク）'] == 1, '人工膝関節置換術（高リスク）'] = '人工膝関節置換術（高リスク）'
template_op.loc[template_op['下肢手術（中リスク）'] == 1, '下肢手術（中リスク）'] = '下肢手術（中リスク）'
template_op.loc[template_op['股関節骨折手術（高リスク）。　（大腿骨骨幹部を含む）'] == 1, '股関節骨折手術（高リスク）'] = '股関節骨折手術（高リスク）'
template_op.loc[template_op['上肢手術（中リスク）。（腸骨からの採骨、下肢から。　神経・皮膚の採取を伴う）'] == 1, '上肢手術（中リスク）'] = '上肢手術（中リスク）'
template_op.loc[template_op['下肢悪性腫瘍手術（高リスク）'] == 1, '下肢悪性腫瘍手術（高リスク）'] = '下肢悪性腫瘍手術（高リスク）'
template_op.loc[template_op['骨盤骨切り術（高リスク）'] == 1, '骨盤骨切り術（高リスク）'] = '骨盤骨切り術（高リスク）'
template_op.loc[template_op['骨盤骨折（高リスク）'] == 1, '骨盤骨折（高リスク）'] = '骨盤骨折（高リスク）'
template_op.loc[template_op['脊椎・脊髄損傷（中リスク）'] == 1, '脊椎・脊髄損傷（中リスク）'] = '脊椎・脊髄損傷（中リスク）'
template_op.loc[template_op['重症外傷（多発外傷）（高リスク）'] == 1, '重症外傷（多発外傷）（高リスク）'] = '重症外傷（高リスク）'
template_op.loc[template_op['大腿骨遠位部以下の。単独外傷（中リスク）'] == 1, '大腿骨遠位部以下の。単独外傷（中リスク）'] = '大腿骨遠位部以下の単独外傷（中リスク）'
template_op.loc[template_op['重症熱傷／重症脊損（高リスク）'] == 1, '重症熱傷／重症脊損（高リスク）'] = '重症熱傷／重症脊損（高リスク）'
template_op.loc[template_op['非開頭術（低リスク）'] == 1, '非開頭術（低リスク）'] = '非開頭術（低リスク）'
template_op.loc[template_op['開頭術（中リスク）'] == 1, '開頭術（中リスク）'] = '開頭術（中リスク）'

template_op['45分以内の手術（低リスク）'] = template_op['45分以内の手術（低リスク）'].fillna('@')
template_op['45分を超える手術（中リスク）'] = template_op['45分を超える手術（中リスク）'].fillna('@')
template_op['45分以内の検査・処置（低リスク）'] = template_op['45分以内の検査・処置（低リスク）'].fillna('@')
template_op['45分を超える検査・処置（中リスク）'] = template_op['45分を超える検査・処置（中リスク）'].fillna('@')
template_op['上肢手術（低リスク）'] = template_op['上肢手術（低リスク）'].fillna('@')
template_op['人工股関節置換術（高リスク）'] = template_op['人工股関節置換術（高リスク）'].fillna('@')
template_op['脊椎手術（中リスク）'] = template_op['脊椎手術（中リスク）'].fillna('@')
template_op['人工膝関節置換術（高リスク）'] = template_op['人工膝関節置換術（高リスク）'].fillna('@')
template_op['下肢手術（中リスク）'] = template_op['下肢手術（中リスク）'].fillna('@')
template_op['股関節骨折手術（高リスク）'] = template_op['股関節骨折手術（高リスク）'].fillna('@')
template_op['上肢手術（中リスク）'] = template_op['上肢手術（中リスク）'].fillna('@')
template_op['下肢悪性腫瘍手術（高リスク）'] = template_op['下肢悪性腫瘍手術（高リスク）'].fillna('@')
template_op['骨盤骨切り術（高リスク）'] = template_op['骨盤骨切り術（高リスク）'].fillna('@')
template_op['骨盤骨折（高リスク）'] = template_op['骨盤骨折（高リスク）'].fillna('@')
template_op['脊椎・脊髄損傷手術（中リスク）'] = template_op['脊椎・脊髄損傷（中リスク）'].fillna('@')
template_op['重症外傷（多発外傷）（高リスク）'] = template_op['重症外傷（多発外傷）（高リスク）'].fillna('@')
template_op['大腿骨遠位部以下の単独外傷（中リスク）'] = template_op['大腿骨遠位部以下の。単独外傷（中リスク）'].fillna('@')
template_op['重症熱傷／重症脊損（高リスク）'] = template_op['重症熱傷／重症脊損（高リスク）'].fillna('@')
template_op['非開頭術（低リスク）'] = template_op['非開頭術（低リスク）'].fillna('@')
template_op['開頭術（中リスク）'] = template_op['開頭術（中リスク）'].fillna('@')

template_op['術式リスク評価'] = template_op['45分以内の手術（低リスク）'] + ', ' + template_op['45分を超える手術（中リスク）'] + ', ' + template_op[
    '45分以内の検査・処置（低リスク）'] + ', ' + template_op['45分を超える検査・処置（中リスク）'] + ', ' + template_op['上肢手術（低リスク）'] + ', ' + \
                         template_op['人工股関節置換術（高リスク）'] + ', ' + template_op['脊椎手術（中リスク）'] + ', ' + template_op[
                             '人工膝関節置換術（高リスク）'] + ', ' + template_op['下肢手術（中リスク）'] + ', ' + template_op[
                             '股関節骨折手術（高リスク）'] + ', ' + template_op['上肢手術（中リスク）'] + ', ' + template_op[
                             '下肢悪性腫瘍手術（高リスク）'] + ', ' + template_op['骨盤骨切り術（高リスク）'] + ', ' + template_op[
                             '骨盤骨折（高リスク）'] + ', ' + template_op['脊椎・脊髄損傷手術（中リスク）'] + ', ' + template_op[
                             '重症外傷（多発外傷）（高リスク）'] + ', ' + template_op['大腿骨遠位部以下の単独外傷（中リスク）'] + ', ' + template_op[
                             '重症熱傷／重症脊損（高リスク）'] + ', ' + template_op['非開頭術（低リスク）'] + ', ' + template_op['開頭術（中リスク）']
template_op['術式リスク評価'] = template_op['術式リスク評価'].str.replace('@, ', '')
template_op['術式リスク評価'] = template_op['術式リスク評価'].str.replace(', @', '')
template_op['術式リスク評価'] = template_op['術式リスク評価'].str.replace('@', '')

template_op.loc[template_op['リスクなし'] == 1, 'リスクなし'] = 'リスクなし'
template_op.loc[template_op['低リスク'] == 1, '低リスク'] = '低リスク'
template_op.loc[template_op['中リスク'] == 1, '中リスク'] = '中リスク'
template_op.loc[template_op['高リスク'] == 1, '高リスク'] = '高リスク'
template_op.loc[template_op['最高リスク'] == 1, '最高リスク'] = '最高リスク'
template_op['リスクなし'] = template_op['リスクなし'].fillna('@')
template_op['低リスク'] = template_op['低リスク'].fillna('@')
template_op['中リスク'] = template_op['中リスク'].fillna('@')
template_op['高リスク'] = template_op['高リスク'].fillna('@')
template_op['最高リスク'] = template_op['最高リスク'].fillna('@')
template_op['リスク評価記載0'] = template_op['リスクなし'] + ', ' + template_op['低リスク'] + ', ' + template_op['中リスク'] + ', ' + template_op['高リスク'] + ', ' + template_op['最高リスク']
template_op['リスク評価記載0'] = template_op['リスク評価記載0'].str.replace('@, ', '')
template_op['リスク評価記載0'] = template_op['リスク評価記載0'].str.replace(', @', '')
template_op['リスク評価記載0'] = template_op['リスク評価記載0'].str.replace('@', '')
template_op['リスク評価記載'] = (template_op['リスク評価記載0'] == template_op['リスク評価'])

template_op['対策記載'] = template_op[['早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法']].notnull().any(axis=1)

# 監査結果
template_op.loc[template_op['術式リスク評価'] == 'nan', '監査結果0'] = '記載なし'
template_op.loc[template_op['リスク評価記載'] == False, '監査結果1'] = '評価不正確'
template_op.loc[template_op['対策記載'] == False, '監査結果2'] = '対策の記載がない'
template_op['監査結果0'] = template_op['監査結果0'].fillna('@')
template_op['監査結果1'] = template_op['監査結果1'].fillna('@')
template_op['監査結果2'] = template_op['監査結果2'].fillna('@')
template_op['監査結果'] = template_op['監査結果0'] + ', ' + template_op['監査結果1'] + ', ' + template_op['監査結果2']
template_op['監査結果'] = template_op['監査結果'].str.replace('@, ', '')
template_op['監査結果'] = template_op['監査結果'].str.replace(', @', '')
template_op['監査結果'] = template_op['監査結果'].str.replace('@', '')
template_op.loc[template_op['監査結果'] == '', '監査結果'] = 'OK'

# CSVファイルにするための編集
template_op = template_op[['日付', '最終入力者',  '入院診療科', '患者ID', '氏名', '術式リスク評価', 'リスク評価記載', 'リスク評価', '早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法', '対策記載', 'その他', '監査結果']]


# 非手術例のDVTリスク評価
template_others = template_others.iloc[:,
                  [0, 6, 7, 8, 21, 35, 38, 41, 44, 47, 50, 53, 56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, \
                   89, 92, 95, 98, 101, 104, 107, 110, 113, 116, 119, 122, 125, 128, 131, 134, 137, 140, \
                   143, 146, 149, 152, 155, 158, 161, 164, 167, 170, 251]]
template_others = template_others.rename(
    columns={'Unnamed: 113': 'リスク合計点', '診療科': '入院診療科', '可能な範囲で早期離床や積極的な運動を行う': '早期離床', 'ベッド上での運動療法（リハビリテーションを依頼）': 'リハビリ', 'Unnamed: 251': 'その他'})

template_others.loc[template_others['48時間以上の安静臥床（入院前を含む）がないためリスク評価は不要'] == 1, 'リスク評価'] = '評価不要'
template_others.loc[template_others['リスク合計点'] == 0, 'リスク評価'] = 'リスクなし'
template_others.loc[template_others['リスク合計点'] == 1, 'リスク評価'] = '低リスク'
template_others.loc[template_others['リスク合計点'] == 2, 'リスク評価'] = '中リスク'
template_others.loc[template_others['リスク合計点'] == 3, 'リスク評価'] = '中リスク'
template_others.loc[template_others['リスク合計点'] == 4, 'リスク評価'] = '中リスク'
template_others.loc[template_others['リスク合計点'] == 5, 'リスク評価'] = '高リスク'
template_others.loc[template_others['リスク合計点'] == 6, 'リスク評価'] = '高リスク'
template_others.loc[template_others['リスク合計点'] >= 7, 'リスク評価'] = '最高リスク'
template_others['リスク評価'] = template_others['リスク評価'].fillna('記載なし')

# 記載されたリスク評価
template_others.loc[template_others['0点　リスクなし'] == 1, '0点　リスクなし'] = 'リスクなし'
template_others.loc[template_others['1点　低リスク'] == 1, '1点　低リスク'] = '低リスク'
template_others.loc[template_others['2-4点　中リスク'] == 1, '2-4点　中リスク'] = '中リスク'
template_others.loc[template_others['5-6点　高リスク'] == 1, '5-6点　高リスク'] = '高リスク'
template_others.loc[template_others['7点以上　最高リスク'] == 1, '7点以上　最高リスク'] = '最高リスク'
template_others['0点　リスクなし'] = template_others['0点　リスクなし'].fillna('@')
template_others['1点　低リスク'] = template_others['1点　低リスク'].fillna('@')
template_others['2-4点　中リスク'] = template_others['2-4点　中リスク'].fillna('@')
template_others['5-6点　高リスク'] = template_others['5-6点　高リスク'].fillna('@')
template_others['7点以上　最高リスク'] = template_others['7点以上　最高リスク'].fillna('@')
template_others['リスク評価記載'] = template_others['0点　リスクなし'] + ', ' + template_others['1点　低リスク'] + ', ' + template_others['2-4点　中リスク'] + ', ' + template_others['5-6点　高リスク'] + ', ' + template_others['7点以上　最高リスク']
template_others['リスク評価記載'] = template_others['リスク評価記載'].str.replace('@, ', '')
template_others['リスク評価記載'] = template_others['リスク評価記載'].str.replace(', @', '')
template_others['リスク評価記載'] = template_others['リスク評価記載'].str.replace('@', '')
template_others['リスク評価記載_TF'] = (template_others['リスク評価記載'] == template_others['リスク評価'])
template_others.loc[template_others['リスク評価'] == '評価不要', 'リスク評価記載_TF'] = ''

template_others['対策記載'] = template_others[['早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法']].notnull().any(axis=1)
template_others.loc[template_others['リスク評価'] == '評価不要', '対策記載'] = True

# 監査結果
template_others['監査結果0'] = (template_others['リスク評価記載'] == '') & (template_others['リスク評価'] != '評価不要')

template_others.loc[template_others['監査結果0'] == True, '監査結果0'] = 'リスク評価の記載がない'
template_others.loc[template_others['監査結果0'] == False, '監査結果0'] = '@'
template_others.loc[template_others['リスク評価記載_TF'] == False, '監査結果1'] = '評価不正確'
template_others.loc[template_others['対策記載'] == False, '監査結果2'] = '対策の記載がない'
template_others['監査結果0'] = template_others['監査結果0'].fillna('@')
template_others['監査結果1'] = template_others['監査結果1'].fillna('@')
template_others['監査結果2'] = template_others['監査結果2'].fillna('@')
template_others['監査結果'] = template_others['監査結果0'] + ', ' + template_others['監査結果1'] + ', ' + template_others['監査結果2']
template_others['監査結果'] = template_others['監査結果'].str.replace('@, ', '')
template_others['監査結果'] = template_others['監査結果'].str.replace(', @', '')
template_others['監査結果'] = template_others['監査結果'].str.replace('@', '')
template_others.loc[template_others['監査結果'] == '', '監査結果'] = 'OK'
template_others['リスク評価記載'] = template_others['リスク評価記載_TF']

template_others = template_others[
    ['日付', '最終入力者',  '入院診療科', '患者ID', '氏名', 'リスク評価記載', 'リスク評価', '早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法', '対策記載', 'その他', '監査結果']]


# 2つのテンプレートの統合
audit_templates = pd.concat([template_op, template_others])
audit_templates.to_csv(folder_path + 'data/13監査結果/テンプレート統合.csv', encoding='utf-8_sig')


# 患者データとテンプレートの統合
audit = pd.merge(inpatients, audit_templates, on='患者ID', suffixes=['_Pa', '_Te'])
audit.to_csv(folder_path + 'data/13監査結果/テンプレートと患者一覧の統合.csv', encoding='utf-8_sig')

# 重複IDの調整
audit['重複ID'] = audit.duplicated(subset='患者ID', keep=False)
audit['記載科と入院科'] = (audit['入院診療科_Pa'] == audit['入院診療科_Te'])
audit['重複ID_1'] = (audit['重複ID'] == True) & (audit['記載科と入院科'] == False)
audit['監査結果'] = audit[audit['記載科と入院科'] == True]['監査結果']
audit['重複ID_2'] = audit[audit['記載科と入院科'] == False].duplicated(subset='患者ID', keep=False)
audit['重複ID_3'] = audit['重複ID_2'].fillna(False)
audit['重複drop'] = (audit['重複ID_1'] == True) & (audit['重複ID_2'] == False)
audit = audit[audit['重複drop'] == False]
audit = audit.drop(columns=['重複ID', '重複ID_1', '重複ID_2', '重複ID_3', '重複drop'])
audit['監査結果'] = audit['監査結果'].fillna('記載なし')

audit = audit.rename(columns={'日付_Pa': '日付', '氏名_Pa': '氏名', '日付_Te': '記載日', '入院診療科_Pa': '入院診療科', '入院診療科_Te': '記載診療科', '最終入力者': '記載者'})
audit = audit[
    ['登録区分', '日付', '患者ID', '氏名', '性別', '入院診療科', '入院主治医', '記載日', '記載者', '記載診療科', '術式リスク評価', 'リスク評価', 'リスク評価記載', '早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法', '対策記載', 'その他', '監査結果']]

# 記載科と入院診療科が異なる初期アセスメントしかない症例を1つに限定してcsvファイルに書き出す
audit = audit.drop_duplicates(subset='患者ID')
audit.to_csv(folder_path + 'data/13監査結果/全ての静脈血栓塞栓症リスク評価の監査結果.csv', encoding='utf_8_sig')

