# 家計簿のシートへの購入情報の追加と予算等の表示
import pandas as pd # データフレームの作成などに使われるライブラリ(パンダス)
                    # (openpyxl(オープンピーエクセル)というExcelファイル操作に使うモジュールを内包)
import psutil       # 実行プロセスの取得や終了、CPU使用率や通信強度を取得できるライブラリ(ピーエスユーティル)
import math
import codecs       # テキストファイルを操作するのに使われるモジュール(コーデックス)
import os 

# シート名を管理するSheetクラス
class SheetName:
    def __init__(self, name):
        self.name = name     # インスタンス変数に値を代入
    
    def getSheetName(self):
        return self.name
    
    def setSheetName(self, newname):
        self.name = newname


# エントランスsu
print('家計簿プログラムへようこそ！\n\n扱うシートを入力してください')
s = SheetName(str(input('シート名：')))
Sheet = s.getSheetName()

# パス指定
DateFile = './Date/Date' + Sheet[0:4] + '_' + Sheet[5] + '.txt' # 日付を格納するテキストファイル
path = './HhA_Sheet.xlsx'                                       # Excelファイル

# 日付の格納関数
def SaveDT(date):
    Sheet = s.getSheetName()
    DateFile = './Date/Date' + Sheet[0:4] + '_' + Sheet[5] + '.txt' # ノングローバル変数
    datetext = Sheet[5] + '/' + str(date) +'\n'
    with codecs.open(DateFile, 'a', 'utf8') as d:
        d.write(datetext)

# 日付の参照
def LoadDT(k):
    Sheet = s.getSheetName()
    DateFile = './Date/Date' + Sheet[0:4] + '_' + Sheet[5] + '.txt' # ノングローバル変数
    Datelist = []
    with codecs.open(DateFile, 'r', 'utf8') as d:
        Datelist = d.readlines()
    return Datelist[k]
    
# 格納している日付の数の参照
def getLenDT():
    Sheet = s.getSheetName()
    DateFile = './Date/Date' + Sheet[0:4] + '_' + Sheet[5] + '.txt' # ノングローバル変数
    length = 0
    with codecs.open(DateFile, 'r', 'utf8') as d:
        length =  len(d.readlines())
    return length


# 辞書の初期化
taxdic = {int(1):1.00, int(2):1.08, int(3):1.10}
catdic = {int(1):'食事', int(2):'スナック', int(3):'外食', int(4):'部活', int(5):'趣味', int(6):'学業', int(7):'日用品', int(8):'交通', int(9):'ゲーム', int(0):'その他'}

# 商品の追加関数
def addgoods():
    Sheet = s.getSheetName() # グローバル非宣言
    scene = 1
    while scene == 1:
        # Excelファイルの呼び出し
        
        df = pd.read_excel(path, sheet_name = Sheet, index_col = 1)

        # シートの既存データの格納
        length = len(df)
        info = []
        for a in range(length):
            info += [[LoadDT(a), df.iloc[a, 1], df.iloc[a, 2], float(df.iloc[a, 3]), float(df.iloc[a, 4]), float(df.iloc[a, 5]), int(df.iloc[a, 6])]]

        # 購入データの入力
        print('\n< 購入情報の追加 >')
        date  = input('\n何日　　：')                   # 日付
        name  = input('商品名　：')                     # 商品名
        print('\n(1:食事/2:スナック/3:外食/4:部活/5:趣味/6:学業/7:日用品/8:交通/9:ゲーム/0:その他)')
        cat   = catdic[int(input('分類　　：'))]        # 分類
        cost  = float(input('税別価格：'))              # 税別価格
        tnof  = float(input('個数　　：'))              # 個数
        print('\n(1:0%/2:8%/3:10%)')
        tax   = float(taxdic[int(input('税率　　：'))]) # 税率
        total = math.floor(cost * tnof * tax)          # 料金
        SaveDT(date)                                   # 日付の書き込み

        info += [[LoadDT(getLenDT() - 1), name, cat, cost, tnof, tax, total]] # '購入データのリスト'を'既存のデータのリスト'のリストの最後尾に結合

        # データフレームの生成とシートへの書き込み
        with pd.ExcelWriter(path, if_sheet_exists = 'overlay',  engine = 'openpyxl', mode = 'a') as ew:
            dfa = pd.DataFrame(data = info, columns = ['日付', '名前', '分類', '税別価格', '個数', '税率', '料金']) # データフレームの生成
            dfa.to_excel(ew, sheet_name = Sheet)                                    # シートへの書き込み

        # ループの確認
        print('\n商品の追加が終了しました\n次の動作を選択してください\n\n(1:次の商品の追加/0:メニューに戻る)')
        scene = int(input('値：'))

# 出費と予算の出力
def moneyprint():
    # Excelファイルの呼び出し
    df = pd.read_excel(path, sheet_name = s.getSheetName(), index_col = 1)
    length = len(df)

    print('\n< 月の出費と予算 >')
    catsumsum = 0               # 今月の出費合計
    for a in range(10):
        catsum = 0              # 分類別出費
        for b in range(length):
            if df.iloc[b, 2] == catdic[a]:
                    catsum += int(df.iloc[b, 6])
        catsumsum += catsum
        space = ''              # 字揃え空白
        if len(catdic[a]) <= 3:
            space += '　'
            if len(catdic[a]) <= 2:
                space += '　'
        print(str(catdic[a]) + space + '：' + str(catsum))
    print('\n月の出費：' + str(catsumsum))
    print('\n月の予算：' + str((30000 - catsumsum)))

    input('\nEnterを押すとメニューへ戻ります')

# シートの移動
def changeSheet():
    s.setSheetName(str(input('\n移動先のシート：')))
    print('\n現在のシート：' + s.getSheetName())

# シートの生成
def makeSheet():        
    print('\n< シートの追加 >')
    SheetName = str(input('新しいシート名：'))
    with pd.ExcelWriter(path, engine = 'openpyxl', if_sheet_exists = 'overlay', mode = 'a') as ew:
        df = pd.DataFrame()
        df.to_excel(ew, sheet_name = SheetName)
    scene3 = int(input('\nシートの追加が完了しました\n新しいシートへ移動しますか(1:Yes/0:No)\n値：'))
    if scene3 == 1:
        s.setSheetName(SheetName)
        print('\n現在のシート　：' + s.getSheetName())

# Excelファイルを閉じる
def closeExcel(process_name):
    for process in psutil.process_iter(['pid', 'name']):  # 実行中のプロセスを取得(pid：プロセスID、name：プロセス名)
        if process.name() == process_name:                # 探索中のプロセスの名前と引数の一致を確認
            process.kill()                                # プロセスを終了 

# Excelファイルを開く
def openExcel():
    print('\nExcelファイルを開きます\n')
    os.startfile('HhA_sheet.xlsx')
    input('閉じるときはEnterを押してください\n')
    closeExcel('EXCEL.EXE')
    


# 実行本体
scene = 1
while scene == 1:
    scene2 = int(input('\n< メニュー >\n1 : 商品の追加\n2 : 予算表示等\n3 : シートの移動\n4 : シートの追加\n5 : Excelファイルの確認\n0 : 終了\n\n値：'))
    if scene2 == 1:
        addgoods()    # 購入情報の追加
    if scene2 == 2:
        moneyprint()  # 予算と出費の表示
    if scene2 == 3:
        changeSheet() # 作業シートの移動
    if scene2 == 4:
        makeSheet()   # シートの新規作成
    if scene2 == 5:
        openExcel()   # Excelの起動、終了
    if scene2 == 0:
        scene = 0
