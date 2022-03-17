import xlwings as xw
import PySimpleGUI as sg
import os
import time
import math

swV=[]
swI=[]
    
def sendCmd(com,row):
    wbx.sheets["labo_sheet"].activate()
    pg1=wbx.macro('セルごとに送信')
    pg2=wbx.macro('セルごとに送受信')
    xw.Range((row,5)).value=com
    xw.Range((row,5)).select()
    if com[-1]=='?':
        pg2()
    else:
        pg1()
    return

def get_category(cat):
    return category[cat]

def active_category(cat):
    wbx.sheets[actsheet[cat]].activate()
    return

def get_integ(integ):
    return integral[integ]

row=4
col=4
comlist=[]


#発生
valuesA=['発生モード','MD0','MD1','MD2','MD3','MD?',
         '発生ファンクション','VF','IF','V?','I?',
         '発生レンジ','SVRX','SVR3','SVR4','SVR5','SVR?',
         'SIRX','SIR-1','SIR0','SIR1','SIR2','SIR3','SIR4','SIR?',
         '発生値','SOV','SOI','SOV?','SOI?',
         'リミット値','LMV','LMI','LMV?','LMI?',
         'サスペンド電圧','SUV','SUV?',
         'サスペンドHiz/Loz','SUZ0','SUZ1','SUZ?',
         'パルス・ベース値','DBV','DBI','DBV?','DBI?',
         'トリガ・モード','M0','M1','M?',
         'オペレート／スタンバイ','SBY','OPR','SUS','SBY?','OPR?','SUS?',
         'リモート・センシング','RS0','RS1','RS?',
         '時間パラメータ','SP','SP?','SD','SD?',
         'レスポンス','FL0','FL1','FL?']
#スイープ
valuesB=['リニア・スイープ','SN','SN?',
         'フィクスドレベル・スイープ','SF','SF?',
         'ランダム・スイープ','SC','SC?',
         '２ステップ・リニア・スイープ','SM','SM?'
         'スイープタイプ','SX?',
         'ランダム・スイープメモリデータ','N','P','N?','NP?','RSAV','RLOD','RCLR',
         'パルス掃引ベース値','BS','BS?',
         'バイアス値','SB','SB?',
         'ＲＴＢ','RB0','RB1','RB?',
         'スイープレンジ','SR0','SR1','SR?',
         'リバース・モード','SV0','SV1','SV?',
         'スイープリピート回数','SS','SS?',
         'スイープの停止','SWSP',
         'トリガ','*TRG']
#測定
valuesC=['ファンクション','F0','F1','F2','F3','F?',
         '測定レンジ','R0','R1','R?',
         '測定ファンクション連動モード','FX0','FX1','FX?',
         '積分時間','IT0','IT1','IT2','IT3','IT4','IT5','IT6','IT7','IT8','IT?',
         'オート・ゼロ','AZ0','AZ1','AZ?',
         '単位表示切換え','DM0','DM1','DM?',
         '測定表示桁数','RE3','RE4','RE5','RE?',
         '測定オート・レンジ・ディレィ','RD','RD?',
         '測定バッファメモリ','ST0','ST1','ST2','ST?','RL','RN','RN?','SZ?','RNM','RNM?']
#演算
valuesD=['ＮＵＬＬ演算','NL0','NL1','NL?','KNL','KNL?',
         'コンペア演算','CO0','CO1','CO?','KHI','KLO','KHI?','KLO?',
         'スケーリング','SCL0','SCL1','SCL?','KA','KB','KC','KA?','KB?','KC?',
         'ＭＡＸ／ＭＩＮ','MN0','MN1','MN?','AVE?','MAX?','MIN?','TOT?','AVN?']
#システム
valuesE=['ユーザー・パラメータ','STP0','STP1','STP2','STP3','SINI','RCLP0','RCLP1','RCLP2','RCLP3','RINI',
         '機器の初期化','*RST','C',
         '機器情報','*IDN?',
         '電源周波数','LF?',
         '通知ブザー','NZ0','NZ1','NZ?',
         '比較演算結果ブザー','BZ0','BZ1','BZ2','BZ3','BZ4','BZ?',
         'リミット検出ブザー','UZ0','UZ1','UZ?',
         'セルフテスト','*TST?','TER?',
         'エラーログ','ERL?','ERC?',
         'インタロック設定','OP0','OP1','OP2','OP3','OP4','OP?',
         '同期制御信号の入出力設定','CP0','CP1','CP2','CP3','CP4','CP5','CP6','CP?','CW0','CW1','CW?']
#リモート
valuesF=['ブロック・デリミタ','DL0','DL1','DL2','DL3','DL?',
         'ヘッダの出力','OH0','OH1','OH?',
         'ＳＲＱ','S0','S1','S?',
         'ステータス','*STB?','*SRE','*SRE?','*ESR?','*ESE','*ESE?','DSR?','DSE','DSE?','ERR?','*CLS',
         'オペレーション・コンプリート','*OPC','*OPC?','*WAI']
#校正
valuesG=['校正ＳＷ','CAL0','CAL1','CAL?',
         '校正データ','XINI','XWR',
         '校正実行','XVS','XIS','XVLH','XVLL','XILH','XILL','XVM','XIM',
         '校正レンジ','XR-1','XR0','XR1','XR2','XR3','XR4',
         '校正データ','XDAT','XD','XADJ','XUP','XDN','XNXT']

category={'発生':valuesA,
          'スイープ':valuesB,
          '測定':valuesC,
          '演算':valuesD,
          'システム':valuesE,
          'リモート':valuesF,
          '校正':valuesG}

actsheet={'発生':"③-1コマンド設定　発生",
          'スイープ':"③-2コマンド設定　スイープ",
          '測定':"③-3コマンド設定　測定",
          '演算':"③-4コマンド設定　演算",
          'システム':"③-5コマンド設定　システム",
          'リモート':"③-6コマンド設定　リモートと校正",
          '校正':"③-6コマンド設定　リモートと校正"}

integral={'100us':'IT0',
          '500us':'IT1',
          '1ms':'IT2',
          '5ms':'IT3',
          '10ms':'IT4',
          '1PLC':'IT5',
          '100ms':'IT6',
          '200ms':'IT7'}

swkeys=['vw1','vr1','vw2','vr2','vw3','vr3','vw4','vr4','vw5','vr5']

prdkeys=['p1','p2','p3','p4','p5','p6','p7','p8','p9','p10']

swnums={'1':1,
        '2':2,
        '3':3,
        '4':4,
        '5':5,
        '6':6,
        '7':7,
        '8':8,
        '9':9,
        '10':10}

sg.theme('DarkTeal7')
          
layout1= [
    [sg.Text('コマンド直接入力')],
    [sg.InputText('', size=(10, 1), key='text'),sg.Button('OK',key='textok')],
    [sg.Text('コマンドのカテゴリ選択')],
    [sg.Combo(values=['発生','スイープ','測定','演算','システム','リモート','校正'],
              size=(10, 1),key='selectCategory',enable_events=True)],
    [sg.Listbox(values='',size=(35,8),key='selectCommand',enable_events=True)],
    [sg.Text('数値入力')],
    [sg.InputText('', size=(10, 1), key='number',enable_events=True),sg.Button('OK',key='numok')],
    [sg.Button('個別に送信',key='sendonce'),sg.Button('すべて送信',key='sendall')
     ,sg.Button('1つクリア',key='stack'),sg.Button('すべてクリア',key='clear')],
    [sg.Output(size=(80,3),key='output')],
    [sg.Text('コマンドセットの保存/読み込み')],
    [sg.InputText('保存先を入力', size=(20, 1), key='save'),sg.Button('保存',key='saveok')],
    [sg.FileBrowse('読み込むファイルを選択', key='load',enable_events=True)]
]

layout2= [
    [sg.Text('パラメータを入力してください。')],
    [sg.Text('')],
    [sg.Text('スタート値'),sg.InputText('0', size=(10, 1), key='startval'),sg.Text('A'),
     sg.Text('    ストップ値'),sg.InputText('5', size=(10, 1), key='stopval'),sg.Text('mA')],
    [sg.Text('ステップ値'),sg.InputText('10', size=(10, 1), key='stepval'),sg.Text('μA')],
    [sg.Text('ホールド時間'),sg.InputText('1', size=(10, 1), key='holdtime'),sg.Text('ms'),
     sg.Text('    ソース・ディレイ時間'),sg.InputText('0.03', size=(10, 1), key='srcdelay'),sg.Text('ms')],
    [sg.Text('メジャー・ディレイ時間'),sg.InputText('0.05', size=(10, 1), key='measdelay'),
     sg.Text('ms'),sg.Text('    パルス幅'),sg.InputText('0.05', size=(10, 1), key='plswidth'),sg.Text('ms')],
    [sg.Text('ピリオド'),sg.InputText('0.5', size=(10, 1), key='period'),sg.Text('ms')],
    [sg.Text('繰り返し回数'),sg.InputText('5', size=(10, 1), key='ddcycle')],
    [sg.Button('測定開始',key='sweep')]
    #[sg.Button('測定停止',key='measstop')]
]    

layout3= [
    [sg.Text('パラメータを入力してください。')],
    [sg.Text('')],
    [sg.Text('スタート値'),sg.InputText('-1', size=(10, 1), key='ivstart'),sg.Text('V'),sg.Text('    ミドル値'),sg.InputText('0', size=(10, 1), key='ivmiddle'),sg.Text('V')],
    [sg.Text('ストップ値'),sg.InputText('1', size=(10, 1), key='ivstop'),sg.Text('V')],
    [sg.Text('第1ステップ値'),sg.InputText('0.01', size=(10, 1), key='ivstepi'),sg.Text('V'),sg.Text('    第2ステップ値'),sg.InputText('0.01', size=(10, 1), key='ivstepii'),sg.Text('V')],
    [sg.Text('ホールド時間'),sg.InputText('3', size=(10, 1), key='ivhold'),sg.Text('ms'),sg.Text('    ソース・ディレイ時間'),sg.InputText('0.03', size=(10, 1), key='ivsd'),sg.Text('ms')],
    [sg.Text('メジャー・ディレイ時間'),sg.InputText('4', size=(10, 1), key='ivmd'),sg.Text('ms'),sg.Text('    電流リミッタ値'),sg.InputText('0.1', size=(10, 1), key='ivlm'),sg.Text('A')],
    [sg.Text('ピリオド'),sg.InputText('20', size=(10, 1), key='ivperiod'),sg.Text('ms'),sg.Text('    バイアス値'),sg.InputText('0', size=(10, 1), key='ivvias'),sg.Text('V')],
    [sg.Text('繰り返し回数'),sg.InputText('1', size=(10, 1), key='ivivcycle'),sg.Checkbox('リバース',key='isreverse',default=False)],
    [sg.Button('測定開始',key='ivsweep')]
    #[sg.Button('測定停止',key='measstop')]
]

layout4= [
    [sg.Text('パラメータを入力してください。')],
    [sg.Text('')],
    [sg.Text('ステップ数'),sg.Combo(values=['1','2','3','4','5','6','7','8','9','10'],size=(10, 1),key='swnum',default_value='4',enable_events=True)],
    [sg.Text('V1'),sg.InputText('3', size=(5, 1), key='vw1',disabled=False),sg.Text('V'),sg.Text('    P1'),sg.InputText('50', size=(5, 1), key='p1',disabled=False),sg.Text('ms'),sg.Text('    V2'),sg.InputText('0', size=(5, 1), key='vr1',disabled=False),sg.Text('V'),sg.Text('    P2'),sg.InputText('5', size=(5, 1), key='p2',disabled=False),sg.Text('ms')],
    [sg.Text('V3'),sg.InputText('0.02', size=(5, 1), key='vw2',disabled=False),sg.Text('V'),sg.Text('    P3'),sg.InputText('20', size=(5, 1), key='p3',disabled=False),sg.Text('ms'),sg.Text('    V4'),sg.InputText('-1.5', size=(5, 1), key='vr2',disabled=False),sg.Text('V'),sg.Text('    P4'),sg.InputText('50', size=(5, 1), key='p4',disabled=False),sg.Text('ms')],
    [sg.Text('V5'),sg.InputText('0', size=(5, 1), key='vw3',disabled=True),sg.Text('V'),sg.Text('    P5'),sg.InputText('5', size=(5, 1), key='p5',disabled=True),sg.Text('ms'),sg.Text('    V6'),sg.InputText('0.02', size=(5, 1), key='vr3',disabled=True),sg.Text('V'),sg.Text('    P6'),sg.InputText('20', size=(5, 1), key='p6',disabled=True),sg.Text('ms')],
    [sg.Text('V7'),sg.InputText('-0.5', size=(5, 1), key='vw4',disabled=True),sg.Text('V'),sg.Text('    P7'),sg.InputText('50', size=(5, 1), key='p7',disabled=True),sg.Text('ms'),sg.Text('    V6'),sg.InputText('0.02', size=(5, 1), key='vr4',disabled=True),sg.Text('V'),sg.Text('    P8'),sg.InputText('50', size=(5, 1), key='p8',disabled=True),sg.Text('ms')],
    [sg.Text('V9'),sg.InputText('0.8', size=(5, 1), key='vw5',disabled=True),sg.Text('V'),sg.Text('    P9'),sg.InputText('50', size=(5, 1), key='p9',disabled=True),sg.Text('ms'),sg.Text('    V10'),sg.InputText('0.02', size=(5, 1), key='vr5',disabled=True),sg.Text('V'),sg.Text('    P10'),sg.InputText('50', size=(5, 1), key='p10',disabled=True),sg.Text('ms')],
    [sg.Text('ホールド時間'),sg.InputText('3', size=(10, 1), key='plshold'),sg.Text('ms'),sg.Text('    ソース・ディレイ時間'),sg.InputText('0.03', size=(10, 1), key='plssd'),sg.Text('ms')],
    [sg.Text('メジャー・ディレイ時間'),sg.InputText('4', size=(10, 1), key='plsmd'),sg.Text('ms')],
    [sg.Text('電流リミッタ値(上)'),sg.InputText('3', size=(10, 1), key='plslmu'),sg.Text('mA'),sg.Text('    電流リミッタ値(下)'),sg.InputText('-3', size=(10, 1), key='plslmd'),sg.Text('mA')],
    [sg.Text('積分時間'),sg.Combo(values=['100us','500us','1ms','5ms','10ms','1PLC','100ms','200ms'],size=(10, 1),key='integ',default_value='10ms',enable_events=True)],
    [sg.Text('繰り返し回数'),sg.InputText('5', size=(10, 1), key='cycle')],
    [sg.Button('測定開始',key='plsmeasure2')]
    #[sg.Button('測定停止',key='measstop')]
]    

main_layout= [
    [sg.TabGroup([[sg.Tab('コマンド入力',layout1),sg.Tab('電解水電池',layout2),sg.Tab('I-V測定',layout3),sg.Tab('スイッチング',layout4)]])]
]

window = sg.Window('6241 CommandSender ver.2.0',size=(600,600)).Layout(main_layout)

wbx=xw.Book('labo.xlsm')
wbx.sheets["labo_sheet"].activate()
for i in range(4,24):
    xw.Range((i,5)).value=('')
    xw.Range((i,6)).value=('')
    xw.Range((i-1,7)).value=('')
    xw.Range((i,8)).value=('')

while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED: 
        break

    if event == 'sendonce': 
        sendCmd(comlist[-1],row)
        row+=1

    if event == 'sendall':
        for i in range(len(comlist)):
            sendCmd(comlist[i],row)
            row+=1
            time.sleep(0.003)

    if event == 'selectCategory':
        catt=values['selectCategory']
        selcat=get_category(catt)
        active_category(catt)
        window['selectCommand'].Update(values=selcat)

    if event == 'selectCommand':
        window['number'].Update('')
        comlist.append(values['selectCommand'][0])
        window['output'].Update('')
        print(comlist)

    if event == 'stack':
        comlist.pop(-1)
        window['output'].Update('')
        print(comlist)

    if event == 'clear':
        comlist=[]
        window['output'].Update('')
        print(comlist)

    if event == 'numok':
        last=comlist[-1] + str(values['number'])
        comlist[-1]=last
        window['output'].Update('')
        print(comlist)

    if event == 'textok':
        comlist.append(values['text'])
        window['output'].Update('')
        print(comlist)
        
    if event == 'saveok':
        f=open(values['save']+'.txt','w')
        for i in range(len(comlist)):
            f.write(comlist[i]+'\n')
        f.close()

    if event == 'load':
        f=open(values['load'],'r')
        loadcomlist=f.readlines()
        for i in range(len(loadcomlist)):
            loadcomlist[i]=loadcomlist[i].rstrip('\n')
        f.close()
        comlist=loadcomlist
        window['output'].Update('')
        print(comlist)

    
    if event == 'sweep':
        comlist=[]
        comlist.extend(['MD2','IF','F1','SR1','OH0'])
        comlist.append('SN'+str(float(values['startval']))+','+str(float(values['stopval'])*0.001)+\
                       ','+str(float(values['stepval'])*1e-6))
        comlist.append('SP'+str(float(values['holdtime']))+','+str(float(values['measdelay']))+\
                       ','+str(float(values['period']))+','+str(float(values['plswidth'])))
        comlist.append('SD'+str(float(values['srcdelay'])))
        window['output'].Update('')
        print(comlist)
        dcycle=((abs(float(values['startval']))-abs(float(values['stopval'])*0.001))/(float(values['stepval'])*1e-6))
        xw.Range((2,5)).value=((math.floor(dcycle)+1)*int(values['ddcycle']))
        for i in range(len(comlist)):
            sendCmd(comlist[i],row)
            
            time.sleep(0.003)
        pg3=wbx.macro('トリガ測定と格納')
        pg3()

    if event == 'ivsweep':
        comlist=[]
        comlist.extend(['MD2','VF','F2','SR1','OH0'])
        comlist.append('SM'+str(float(values['ivstart']))+','+str(float(values['ivmiddle']))+\
                       ','+str(float(values['ivstop']))+','+str(float(values['ivstepi']))+\
                       ','+str(float(values['ivstepii'])))
        comlist.append('SB'+str(float(values['ivvias'])))
        comlist.append('LMI'+str(float(values['ivlm'])))
        comlist.append('SP'+str(float(values['ivhold']))+','+str(float(values['ivmd']))+\
                       ','+str(float(values['ivperiod'])))
        comlist.append('SD'+str(float(values['ivsd'])))
        comlist.extend(['IT4','AZ0','ST1','RS1'])
        ivcycle=((abs(float(values['ivstart']))-abs(float(values['ivmiddle'])))/(float(values['ivstepi'])))+((abs(float(values['ivstop']))-abs(float(values['ivmiddle'])))/(float(values['ivstepii'])))
        if values['isreverse']==True:
            rev=2
            comlist.append('SV1')
        else:
            rev=1
            comlist.append('SV0')
        window['output'].Update('')
        print(comlist)
        xw.Range((2,5)).value=((math.floor(ivcycle)+1)*int(values['ivivcycle'])*rev)
        for i in range(len(comlist)):
            sendCmd(comlist[i],row)
            
            time.sleep(0.003)
        pg4=wbx.macro('トリガ測定と格納')
        pg4()

    if event == 'swnum':
        for j in range(10):
            window[swkeys[j]].update(disabled=True)
            window[prdkeys[j]].update(disabled=True)
        for i in range(swnums[values['swnum']]):
            window[swkeys[i]].update(disabled=False)
            window[prdkeys[i]].update(disabled=False)

    if event == 'plsmeasure2':
        for k in range(swnums[values['swnum']]+10):
            xw.Range((3+k,7)).value=('')
            xw.Range((3+k,8)).value=('')
        xw.Range((2,5)).value=(values['cycle'])
        xw.Range((1,8)).value=(values['plshold'])
        xw.Range((2,8)).value=(values['plsmd'])
        for j in range(swnums[values['swnum']]):
            xw.Range((3+j,7)).value=(values[swkeys[j]])
            xw.Range((3+j,8)).value=(values[prdkeys[j]])
        comlist=[]
        comlist.extend(['MD2','VF','F2','SV1','SR1','OH0'])
        comlist.append('SD'+str(float(values['plssd'])))
        comlist.append('LMI'+str(float(values['plslmu'])*0.001)+','+str(float(values['plslmd'])*0.001))
        comlist.append(get_integ(values['integ']))
        comlist.extend(['AZ0','ST1','RS1'])
        window['output'].Update('')
        print(comlist)
        for i in range(len(comlist)):
            sendCmd(comlist[i],row)
            time.sleep(0.003)
        pg6=wbx.macro('トリガ測定と格納sw')
        pg6()

    if event == 'measstop':
        pg10=wbx.macro('トリガ測定と格納をストップする')
        pg10()
        
        

    

window.close()
