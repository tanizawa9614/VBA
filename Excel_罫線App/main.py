import tkinter as tk
from win32com.client import DispatchWithEvents, gencache
import xml.etree.ElementTree as ET

# https://stackoverflow.com/questions/49353023/passing-parameter-to-excels-range-valuerangevaluedatatype-in-win32com
# https://docs.python.org/ja/3/library/xml.etree.elementtree.html#xml.etree.ElementTree.XMLPullParser
# https://qiita.com/feo52/items/150745ae0cc17cb5c866

# 定数の定義
XL_RANGE_VALUE_DEFAULT = 10
XL_RANGE_VALUE_XML_SPREADSHEET = 11
CANVAS_HEIGHT = 40
CANVAS_WIDTH = 60
LINE_HEIGHT = 38
LINE_WIDTH = 58
LINE_COLOR = 'black'
NAME_SPACE = '{urn:schemas-microsoft-com:office:spreadsheet}'

# 監視停止ボタンのコールバック関数
def stop_monitoring():
    root.quit()

# Excelイベントをハンドルするクラス
class ExcelEvents:
    def OnSelectionChange(self, Target):
        update_grid(Target)
    def OnChange(self, Target):
        update_grid(ws.Selection)

# Excelアプリケーションを取得し、新規ブックを作成
xl_app = gencache.EnsureDispatch("Excel.Application")
xl_app.Visible = True
wb = xl_app.Workbooks.Add()
ws = wb.Worksheets(1)

# イベントハンドラを設定
xl_events = DispatchWithEvents(ws, ExcelEvents)

# tkinter GUIのセットアップ
root = tk.Tk()
root.title("Excelセル罫線視覚化ツール")

# グリッドのセルを表すキャンバスを保持する辞書
grid_canvases = {}

# グリッドのセットアップ
def create_grid(center_cell):
    # 中心セルの行と列を取得
    center_row = center_cell.Row
    center_col = center_cell.Column

    print('Address = ' + center_cell.Address)
    if center_row == 1 and center_col == 1:
        center_row = 3
        center_col = 3
    elif center_row == 1 and center_col == 2:
        center_row = 3
        center_col = 4
    elif center_row == 2 and center_col == 1:
        center_row = 4
        center_col = 3
    elif center_row == 2 and center_col == 2:
        center_row = 4
        center_col = 4
    elif center_row == 1 and center_col > 2:
        center_row +=2
        center_col +=2
    elif center_row > 2 and center_col == 1:
        center_row +=2
        center_col +=2
    elif center_row == 2 and center_col > 2:
        center_col +=2
        center_col +=2
    elif center_row > 2 and center_col == 2:
        center_row +=2
        center_col +=2

    # グリッドの各セルを作成
    for i in range(-2, 3):
        for j in range(-2, 3):
            row = center_row + i
            col = center_col + j
            if row > 0 and col > 0:
                cell_id = ws.Cells(row, col).Address
                canvas = tk.Canvas(root, width=CANVAS_WIDTH, height=CANVAS_HEIGHT, bg="white",borderwidth=0, highlightthickness=0)
                canvas.grid(row=i+2, column=j+2, padx=5, pady=5)
                canvas.create_text(30, 20, text=cell_id)
                grid_canvases[cell_id] = canvas
                # キャンバスにクリックイベントをバインド
                canvas.bind("<Button-1>", lambda event, c=canvas, r=row, cl=col: on_canvas_click(event, c, r, cl))

    # borderwidth=0, highlightthickness=0について
    # この記述がないと、左罫線と上罫線がcanvasと重なり非表示になる。
    # 左罫線と上罫線が表示できると、逆に右罫線と下罫線が隠れる
    # よって、LINE_HEIGHTとLINE_WIDTHで若干内側に表示する微調整を行なっている。

# グリッドの更新関数
def update_grid(selected_cell):
    # グリッドを再作成
    create_grid(selected_cell)
    # 罫線情報を更新
    for cell_id, canvas in grid_canvases.items():
        # 罫線を描画
        draw_borders(ws.Range(cell_id), canvas)

# 罫線を描画する関数
def draw_borders(cell, canvas):
    # キャンバスの既存の罫線をクリア
    canvas.delete("border")

    # XML形式のデータを取得
    xml_data = cell.GetValue(XL_RANGE_VALUE_XML_SPREADSHEET)

    # XMLを解析 
    xml_root = ET.fromstring(xml_data)

    # 存在する罫線のPositionをリストに格納
    positions = []
    for border in xml_root.iter(NAME_SPACE + 'Border'):
        positions.append(border.get(NAME_SPACE + 'Position'))

    # 罫線情報に基づいて新しい罫線を描画
    borders = cell.Borders
    if 'Left' in positions:  # 左罫線
        canvas.create_line(1, 0, 1, CANVAS_HEIGHT, tags="border", fill=LINE_COLOR)
    if 'Top' in positions:   # 上罫線
        canvas.create_line(0, 0, LINE_WIDTH, 0, tags="border", fill=LINE_COLOR)
    if 'Right' in positions:   # 右罫線
        canvas.create_line(LINE_WIDTH, 0, LINE_WIDTH, LINE_HEIGHT, tags="border", fill=LINE_COLOR)
    if 'Bottom' in positions:   # 下罫線
        canvas.create_line(0, LINE_HEIGHT, LINE_WIDTH, LINE_HEIGHT, tags="border", fill=LINE_COLOR)
    if 'DiagonalLeft' in positions:  # 左上から右下への対角線
        canvas.create_line(0, 0, LINE_WIDTH, LINE_HEIGHT, tags="border", fill=LINE_COLOR)
    if 'DiagonalRight' in positions:  # 右上から左下への対角線
        canvas.create_line(LINE_WIDTH, 0, 0, LINE_HEIGHT, tags="border",  fill=LINE_COLOR)
    # 監視停止ボタン設置
    stop_button = tk.Button(root, text="監視停止", command=stop_monitoring)
    stop_button.grid(row=6, column=0, columnspan=5)

# キャンバスのクリックイベントハンドラ
def on_canvas_click(event, canvas, row, col):
    # クリックされた位置からExcelのセルを特定
    cell = ws.Cells(row, col)
    x = event.x
    y = event.y
    # クリックされた位置がセルのどの辺かを判定
    if x < 15:  # 左辺をクリック
        toggle_border(cell, 'Left')
    elif x > 45:  # 右辺をクリック
        toggle_border(cell, 'Right')
    elif y < 15:  # 上辺をクリック
        toggle_border(cell, 'Top')
    elif y > 25:  # 下辺をクリック
        toggle_border(cell, 'Bottom')
    elif x < y + 10 and x + y < LINE_HEIGHT + 10:  # 左上から右下への対角線
        toggle_border(cell, 'DiagonalLeft')
    elif x > y - 10 and x + y > LINE_HEIGHT - 10:  # 右上から左下への対角線
        toggle_border(cell, 'DiagonalRight')
    # キャンバス上の罫線を再描画
    draw_borders(cell, canvas)

# 罫線の切り替え関数
# (解析したXMLは要素の検索に使用。検索結果を使いながら元のXMLデータを文字列処理)
def toggle_border(cell, side):
    # XML形式のデータを取得
    xml_data = cell.GetValue(XL_RANGE_VALUE_XML_SPREADSHEET)

    # XMLを解析 
    xml_root = ET.fromstring(xml_data)

    # 存在する罫線のPositionをリストに格納
    positions = []
    for border in xml_root.iter(NAME_SPACE + 'Border'):
        positions.append(border.get(NAME_SPACE + 'Position'))
    
    # Cell要素の有無(セルに値をセットしたことがあるか）調べるために取得
    StyleIDs = []
    for border in xml_root.iter(NAME_SPACE + 'Cell'):
        StyleIDs.append(border.get(NAME_SPACE + 'StyleID'))

    # 罫線を引く/削除する
    # 一度も書式をセットしたことのないセル = positionsが空かつ</Borders>(閉じタグ)がない
    if len(positions) == 0 and len(StyleIDs) == 0:
        print('一度も書式をセットしたことがないセル')
        # Font.Colorを一旦赤→黒に戻すことで、DefaultではないStyleタグを作成する
        cell.Font.Color = int("FF0000", 16)
        cell.Font.Color = int("000000", 16)
        
        # XML形式のデータを再取得
        xml_data = cell.GetValue(XL_RANGE_VALUE_XML_SPREADSHEET)

        # XMLを再度解析 
        xml_root = ET.fromstring(xml_data)

        # Defaultではない方のStyleタグを取得
        style_tag = getStyleTag(xml_root)

        text_to_append = '\n<Borders>\n    <Border ss:Position="' + side +'" ss:LineStyle="Continuous" ss:Weight="1"/>\n   </Borders>' 
        style_tag_end_positon = xml_data.find(style_tag) + len(style_tag)
        xml_data = xml_data[:style_tag_end_positon] + text_to_append + xml_data[style_tag_end_positon:]
        print(xml_data)
        # 更新後のXML形式のデータを反映
        cell.SetValue(XL_RANGE_VALUE_XML_SPREADSHEET,xml_data)

    # 罫線をセットしたことがあるが、現在罫線が1本もないセルに罫線を追加
    elif len(positions) == 0 and len(StyleIDs) > 0:
        print('罫線をセットしたことがあるが、現在罫線が1本もないセルに罫線を追加')
        # Defaultではない方のStyleタグを取得
        style_tag = getStyleTag(xml_root)
        style_tag_end_positon = xml_data.find(style_tag) + len(style_tag)
        before_style_tag = xml_data[:style_tag_end_positon]
        after_style_tage  = xml_data[style_tag_end_positon:]

        text_to_replace = '\n<Borders>\n    <Border ss:Position="' + side +'" ss:LineStyle="Continuous" ss:Weight="1"/>\n   </Borders>' 
        after_style_tage = after_style_tage.replace('<Borders/>',text_to_replace)
        xml_data = before_style_tag + after_style_tage

        print(xml_data)

        # 更新後のXML形式のデータを反映
        cell.SetValue(XL_RANGE_VALUE_XML_SPREADSHEET,xml_data)

    # 罫線を削除する場合 = positionsにsideが含まれている
    elif side in positions:
        print('罫線を削除する場合')
        text_to_remove = '<Border ss:Position="' + side +'" ss:LineStyle="Continuous" ss:Weight="1"/>' 
        xml_data = xml_data.replace(text_to_remove,'')
        print(xml_data)
        # 更新後のXML形式のデータを反映
        cell.SetValue(XL_RANGE_VALUE_XML_SPREADSHEET,xml_data)
    
    # 罫線が1本以上あるセルに別の罫線を追加
    else:
        print('罫線が1歩以上あるセルに別の罫線を追加')
        text_to_append = '<Border ss:Position="' + side +'" ss:LineStyle="Continuous" ss:Weight="1"/>' 
        borders_end_positon = xml_data.find('</Borders>')
        xml_data = xml_data[:borders_end_positon] + text_to_append + xml_data[borders_end_positon:]
        print(xml_data)
        # 更新後のXML形式のデータを反映
        cell.SetValue(XL_RANGE_VALUE_XML_SPREADSHEET,xml_data)

#DefaultではないほうのStyleタグを取得
def getStyleTag(xml_root):
    StyleIDs = []

    for border in xml_root.iter(NAME_SPACE + 'Cell'):
        StyleIDs.append(border.get(NAME_SPACE + 'StyleID'))

    style_id = StyleIDs[len(StyleIDs) - 1]
    style_tag = '<Style ss:ID="' + style_id+'">'

    return style_tag

# GUIを実行
root.mainloop()

# Excelを閉じる
wb.Close(False)
xl_app.Quit()
