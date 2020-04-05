import pandas as pd
import tkinter as tk
import tkinter.messagebox as mb
import os

# folder_path
folder_path = './'


def first_gui():
    root = tk.Tk()
    root.title('DVTリスク評価の監査プログラム')
    root.geometry('500x280')
    # first_guiを中央に配置する
    root.attributes("-topmost", True)
    root.update_idletasks()
    ww = root.winfo_screenwidth()
    lw = root.winfo_width()
    wh = root.winfo_screenheight()
    lh = root.winfo_height()
    root.geometry(str(lw) + "x" + str(lh) + "+" + str(int(ww / 2 - lw / 2)) + "+" + str(int(wh / 2 - lh / 2)))
    desc_label1 = tk.Label(text='\n静脈血栓塞栓症リスク評価の監査', font=("MS ゴシック", 17, 'bold'), fg='brown')
    desc_label2 = tk.Label(text='\nカルテ記載の監査を実行します。\n\n電子カルテから「data」フォルダ内の\n「0患者情報」フォルダと「7DVT」フォルダに\nそれぞれCSVファイルを入れて、\n解析開始ボタンを押してください。', font=("MS ゴシック", 14))
    desc_label1.pack()
    desc_label2.pack()

    def file_check():
        if not os.path.isfile(folder_path + 'data/0患者情報/病床管理_入院・転入一覧.csv'):
            if not os.path.isfile(folder_path + 'data/0患者情報/病床管理_転科・担当変更一覧.csv'):
                mb.showerror('Alert', '「0患者情報」フォルダに病床管理_入院・転入一覧.csvと病床管理_転科・担当変更一覧.csvファイルがありません。')
            else:
                mb.showerror('Alert', '「0患者情報」フォルダに病床管理_入院・転入一覧.csvがありません。')
        elif not os.path.isfile(folder_path + 'data/0患者情報/病床管理_転科・担当変更一覧.csv'):
            mb.showinfo('Alert', '「0患者情報」フォルダに病床管理_転科・担当変更一覧.csvファイルがありません。')

        if not os.path.isfile(folder_path + 'data/7DVT/静脈血栓塞栓症リスク評価.csv'):
            if not os.path.isfile(folder_path + 'data/7DVT/非手術例の静脈血栓塞栓症リスク評価.csv'):
                mb.showerror('Alert', '「7DVT」フォルダに静脈血栓塞栓症リスク評価.csvと非手術例の静脈血栓塞栓症リスク評価.csvがありません。')
            else:
                mb.showerror('Alert', '「7DVT」フォルダに静脈血栓塞栓症リスク評価.csvがありません。')
        elif not os.path.isfile(folder_path + 'data/7DVT/非手術例の静脈血栓塞栓症リスク評価.csv'):
            mb.showerror('Alert', '「7DVT」フォルダに非手術例の静脈血栓塞栓症リスク評価.csvがありません。')
        # ファイルが確認できたら、解析に入る
        data_audit()

    # 終了ボタンの表示
    ft_button_end = tk.Button(root, text='終了', bg='#f8b862', font=("", 14), width=6, command=root.destroy)
    ft_button_end.place(x=360, y=230)
    # 解析・ボタンの表示
    ft_button = tk.Button(root, text='解析開始', bg='#cee4ae', font=("", 14), width=20, command=file_check)
    root.bind('<Return>', file_check)
    ft_button.place(x=150, y=200)

    root.mainloop()


def data_audit():
    # import csv files
    inpatients_1 = pd.read_csv(folder_path + 'data/0患者情報/病床管理_入院・転入一覧.csv', skiprows=[0, 1, 2, 3, 4],
                               parse_dates=[4, 26],
                               encoding='cp932')
    inpatients_2 = pd.read_csv(folder_path + 'data/0患者情報/病床管理_転科・担当変更一覧.csv', skiprows=[0, 1, 2, 3],
                               parse_dates=[3, 12],
                               encoding='cp932')
    template_op = pd.read_csv(folder_path + 'data/7DVT/静脈血栓塞栓症リスク評価.csv', skiprows=[0, 1],
                              parse_dates=[0, 1, 2, 22, 23],
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
    audit_admission_periods = ['病床管理_入院・転入一覧'] + admission_date
    transferred_date = start_end_date(transferred_date)
    audit_transferred_periods = ['病床管理_転科・担当変更一覧'] + transferred_date

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

    # 抽出期間のCSVファイルの作成
    audit_periods.to_csv(folder_path + 'data/10監査期間/患者データ・テンプレートの抽出期間.csv', encoding='utf-8_sig')

    # 抽出期間の監査
    audit_periods['開始日'] = pd.to_datetime(audit_periods['開始日'], format="%Y/%m/%d")
    audit_periods['終了日'] = pd.to_datetime(audit_periods['終了日'], format="%Y/%m/%d")

    if audit_periods.loc[2, '開始日'] != audit_periods.loc[0, '開始日']:
        mb.showerror('Warning', '病床管理_入院・転入一覧と静脈血栓塞栓症リスク評価テンプレートの監査期間が異なります。\n「患者データ・テンプレートの抽出期間」ファイルを確認してください。')
    if (audit_periods.loc[2, '終了日'] - audit_periods.loc[0, '終了日']).days > 3:
        mb.showerror('Warning', '静脈血栓塞栓症リスク評価テンプレートの監査期間は病床管理_入院・転入一覧の3日後までとしてください。')

    if audit_periods.loc[3, '開始日'] != audit_periods.loc[0, '開始日']:
        mb.showerror('Warning', '病床管理_入院・転入一覧と非手術例の静脈血栓塞栓症リスク評価テンプレートの監査期間が異なります。\n「患者データ・テンプレートの抽出期間」ファイルを確認してください。')
    if (audit_periods.loc[3, '終了日'] - audit_periods.loc[0, '終了日']).days > 3:
        mb.showerror('Warning', '非手術例の静脈血栓塞栓症リスク評価テンプレートの監査期間は病床管理_入院・転入一覧の3日後までとしてください。')

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
        columns={'Unnamed: 140': '不加的リスク', '診療科': '入院診療科', '可能な範囲で早期離床や積極的な運動を行う': '早期離床',
                 'ベッド上での運動療法（リハビリテーションを依頼）': 'リハビリ', 'Unnamed: 251': 'その他'})

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
    template_op['リスク評価記載0'] = template_op['リスクなし'] + ', ' + template_op['低リスク'] + ', ' + template_op['中リスク'] + ', ' + \
                              template_op['高リスク'] + ', ' + template_op['最高リスク']
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
    template_op['テンプレート監査結果'] = template_op['監査結果0'] + ', ' + template_op['監査結果1'] + ', ' + template_op['監査結果2']
    template_op['テンプレート監査結果'] = template_op['テンプレート監査結果'].str.replace('@, ', '')
    template_op['テンプレート監査結果'] = template_op['テンプレート監査結果'].str.replace(', @', '')
    template_op['テンプレート監査結果'] = template_op['テンプレート監査結果'].str.replace('@', '')
    template_op.loc[template_op['テンプレート監査結果'] == '', 'テンプレート監査結果'] = 'OK'

    # CSVファイルにするための編集
    template_op = template_op[
        ['日付', '最終入力者', '入院診療科', '患者ID', '氏名', '術式リスク評価', 'リスク評価記載', 'リスク評価', '早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング',
         '抗凝固療法', '対策記載', 'その他', 'テンプレート監査結果']]

    # 非手術例のDVTリスク評価
    template_others = template_others.iloc[:,
                      [0, 6, 7, 8, 21, 35, 38, 41, 44, 47, 50, 53, 56, 59, 62, 65, 68, 71, 74, 77, 80, 83, 86, \
                       89, 92, 95, 98, 101, 104, 107, 110, 113, 116, 119, 122, 125, 128, 131, 134, 137, 140, \
                       143, 146, 149, 152, 155, 158, 161, 164, 167, 170, 251]]
    template_others = template_others.rename(
        columns={'Unnamed: 113': 'リスク合計点', '診療科': '入院診療科', '可能な範囲で早期離床や積極的な運動を行う': '早期離床',
                 'ベッド上での運動療法（リハビリテーションを依頼）': 'リハビリ', 'Unnamed: 251': 'その他'})

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
    template_others['リスク評価記載'] = template_others['0点　リスクなし'] + ', ' + template_others['1点　低リスク'] + ', ' + \
                                 template_others['2-4点　中リスク'] + ', ' + template_others['5-6点　高リスク'] + ', ' + \
                                 template_others['7点以上　最高リスク']
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
    template_others['テンプレート監査結果'] = template_others['監査結果0'] + ', ' + template_others['監査結果1'] + ', ' + template_others[
        '監査結果2']
    template_others['テンプレート監査結果'] = template_others['テンプレート監査結果'].str.replace('@, ', '')
    template_others['テンプレート監査結果'] = template_others['テンプレート監査結果'].str.replace(', @', '')
    template_others['テンプレート監査結果'] = template_others['テンプレート監査結果'].str.replace('@', '')
    template_others.loc[template_others['テンプレート監査結果'] == '', 'テンプレート監査結果'] = 'OK'

    template_others['リスク評価記載'] = template_others['リスク評価記載_TF']

    template_others = template_others[
        ['日付', '最終入力者', '入院診療科', '患者ID', '氏名', 'リスク評価記載', 'リスク評価', '早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法',
         '対策記載', 'その他', 'テンプレート監査結果']]


    # 患者データとテンプレートの統合
    audit_templates = pd.concat([template_op, template_others], sort=False)
    audit = pd.merge(inpatients, audit_templates, on='患者ID', suffixes=['_Pa', '_Te'])
    audit = audit.rename(
        columns={'日付_Pa': '日付', '氏名_Pa': '氏名', '日付_Te': '記載日', '入院診療科_Pa': '入院診療科', '入院診療科_Te': '記載診療科',
                 '最終入力者': '記載者'})
    # スリープ科を内科に変更する
    audit = audit.replace('スリープ科', '内科')

    # CSVファイルにする
    audit.to_csv(folder_path + 'data/13監査結果/テンプレートと患者一覧の統合.csv', encoding='utf-8_sig')

    # 重複IDの調整
    audit['重複ID'] = audit.duplicated(subset='患者ID', keep=False)
    audit['記載科と入院科'] = (audit['入院診療科'] == audit['記載診療科'])
    audit['重複ID_1'] = (audit['重複ID'] == True) & (audit['記載科と入院科'] == False)
    audit['監査結果'] = audit[audit['記載科と入院科'] == True]['テンプレート監査結果']
    audit['重複ID_2'] = audit[audit['記載科と入院科'] == False].duplicated(subset='患者ID', keep=False)
    audit['重複ID_3'] = audit['重複ID_2'].fillna(False)
    audit['重複drop'] = (audit['重複ID_1'] == True) & (audit['重複ID_2'] == False)
    audit = audit[audit['重複drop'] == False]
    audit.loc[audit['重複ID'] == True, '要確認_重複ID'] = 'IDの重複'
    audit['要確認_重複ID'] = audit['要確認_重複ID'].fillna('@')
    audit = audit.drop(columns=['重複ID', '重複ID_1', '重複ID_2', '重複ID_3', '重複drop'])

    # 記入日が基準日翌日までに記載されているかを監査
    audit['要確認_記載日'] = ''
    audit['記載日'] = pd.to_datetime(audit['記載日'], format="%Y/%m/%d")
    audit['日付'] = pd.to_datetime(audit['日付'], format='%Y/%m/%d')
    audit['記載日-日付'] = (audit['記載日'] - audit['日付']).apply(lambda x: x.days)
    audit.loc[audit['記載日-日付'] < 0, '要確認_記載日'] = audit['要確認_記載日'].apply(lambda x: x.join('入院・転科日前の記載'))
    audit.loc[audit['記載日-日付'] > 4, '要確認_記載日'] = audit['要確認_記載日'].apply(lambda x: x.join('入院・転科後3日目以降の記載'))
    audit.loc[audit['記載日-日付'] == 0, '要確認_記載日'] = '@'
    audit.loc[audit['要確認_記載日'] == '', '要確認_記載日'] = '@'
    audit.loc[audit['要確認_記載日'] != '@', '監査結果'] = '記載なし'

    # 記載診療科と入院診療科が異なるとき
    audit['記載診療科'] = audit['記載診療科'].fillna('@')
    audit['監査: 記載あり'] = audit['記載診療科'] != '@'
    audit['監査: 診療科不一致'] = (audit['監査: 記載あり'] == True) & (
                audit['記載科と入院科'] == False)
    audit.loc[audit['監査: 診療科不一致'] == True, '要確認_診療科不一致'] = '診療科の不一致'
    audit['要確認_診療科不一致'] = audit['要確認_診療科不一致'].fillna('@')
    audit.loc[audit['要確認_診療科不一致'] != '@', '監査結果'] = '記載なし'

    # その他の記載
    audit['要確認_その他記載'] = audit['その他'].notnull()
    audit.loc[audit['要確認_その他記載'] == False, '要確認_その他記載'] = '@'
    audit.loc[audit['要確認_その他記載'] == True, '要確認_その他記載'] = 'その他の記載を確認'

    # 要確認の集約
    audit['要確認_記載日'] = audit['要確認_記載日'].fillna('@')
    audit['要確認_重複ID'] = audit['要確認_重複ID'].fillna('@')
    audit['要確認_診療科不一致'] = audit['要確認_診療科不一致'].fillna('@')
    audit['要確認_その他記載'] = audit['要確認_その他記載'].fillna('@')
    audit['要確認'] = audit['要確認_記載日'] + ', ' + audit['要確認_重複ID'] + ', ' + audit['要確認_診療科不一致'] + ', ' + audit['要確認_その他記載']
    audit['要確認'] = audit['要確認'].str.replace('@, ', '')
    audit['要確認'] = audit['要確認'].str.replace(', @', '')
    audit['要確認'] = audit['要確認'].str.replace('@', '')

    audit = audit[
        ['登録区分', '日付', '患者ID', '氏名', '性別', '入院診療科', '入院主治医', '記載日', '記載者', '記載診療科', '術式リスク評価', 'リスク評価', 'リスク評価記載',
         '早期離床', 'リハビリ', '間欠的空気圧迫法', '弾性ストッキング', '抗凝固療法', '対策記載', 'その他', 'テンプレート監査結果', '監査結果', '要確認']]

    # 記載科と入院診療科が異なるテンプレート記載しかない症例を1つに限定してcsvファイルに書き出す
    audit = audit.drop_duplicates(subset='患者ID')
    audit.to_csv(folder_path + 'data/13監査結果/全ての静脈血栓塞栓症リスク評価の監査結果.csv', encoding='utf_8_sig')

    mb.showinfo('Finished !', "監査完了")
    ret = mb.askyesno('監査結果の解析', '解析を追加しますか？')
    if ret:
        data_analysis()
    else:
        mb.showinfo('End', 'プログラムを終了します。')


def data_analysis():
    audit = pd.read_csv(folder_path + 'data/13監査結果/全ての静脈血栓塞栓症リスク評価の監査結果.csv', encoding='utf_8_sig')
    audit_crosstab1 = pd.crosstab(index=audit['入院診療科'], columns=audit['監査結果'], margins_name='合計', margins=True)
    audit_crosstab2 = pd.crosstab(index=audit['入院診療科'], columns=audit['監査結果'], margins_name='合計', margins=True, normalize='index')
    audit_crosstab2 = audit_crosstab2.applymap(lambda x: x * 100)
    audit_crosstab1 = audit_crosstab1.rename(columns={'OK': '記載あり'})
    audit_crosstab1 = audit_crosstab1[['記載あり', '評価不正確', '記載なし', '合計']]
    audit_crosstab2 = audit_crosstab2.rename(columns={'OK': '記載あり（％）', '評価不正確': '評価不正確（％）', '記載なし': '記載なし（％）', '合計': '合計（％）'})
    audit_crosstab2 = audit_crosstab2[['記載あり（％）', '評価不正確（％）', '記載なし（％）']]

    audit_crosstab = pd.concat([audit_crosstab1, audit_crosstab2], axis=1)
    dep_number = {'総合診療科': 1, '内科': 2, '循環器内科': 3, '消化器内科': 4, '呼吸器内科': 5, '神経内科': 6, '腎臓内科': 7,
                  'リウマチ・膠原病科': 8, '糖尿病・代謝内科': 9, '血液内科': 10, '緩和ケア内科': 11, '専攻医': 14, '外科': 20,
                  '整形外科': 21, '泌尿器科': 22, '脳神経外科': 23, '眼科': 24, '皮膚科': 25, '耳鼻咽喉科': 26, '麻酔科': 27,
                  '放射線科': 28, 'スリープ科': 29, 'リハビリテーション科': 30, '救急科': 31, '心療内科': 32, '透析科': 33,
                  '婦人科': 34, '乳腺外科': 35, '心臓血管外科': 36, '形成外科': 38, '臨床腫瘍科': 39, '病理診断科': 40,
                  'ICU': 41, '健診科': 45, '歯科口腔外科': 50, '合計': 99}

    audit_crosstab['科番号'] = audit_crosstab.index.map(lambda x: dep_number.get(x))
    audit_crosstab = audit_crosstab.sort_values('科番号')
    audit_crosstab.index.name = '診療科'
    pd.options.display.float_format = '{:.1f}'.format
    audit_crosstab = audit_crosstab.drop(columns='科番号')
    print(audit_crosstab)

    mb.showinfo('End', 'プログラムを終了します。')


first_gui()