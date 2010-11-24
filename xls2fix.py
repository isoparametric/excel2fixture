# -*- coding: utf-8 -*-
import sys
from os import path
from xlrd import open_workbook, xldate_as_tuple, cellname, cellnameabs
import yaml
from argparse import ArgumentParser
import os.path
import datetime
import simplejson as json

class SettingColumn(object):

    def __init__(self, name, type, default):
        self.name = name
        self.type = type
        self.default = default
        pass

class SettingsYaml(object):

    def __init__(self, yaml):

        self.yaml = yaml
        columns = yaml['table']['columns']
        #print columns
        # カラム毎の設定を作っておく
        self.columns = {}
        self.columns['ID'] = {'name':'id', 'column': u'ID', 'type': 'int'}
        for column in columns:
            # カラム名の辞書を生成しておく
            try:
                self.columns[column['column']] = column
            except KeyError:
                print >>sys.stderr, u'name: %sにcolumn要素が足りない' % (column['name'])
                raise
        self.model = yaml['table']['model']

    def is_convert_sheet(self, sheet):
        if self.yaml['table']['sheet'] == sheet.name:
            return True

        return False

    def setting_convert_sheet(self, sheet):

        if self.is_convert_sheet(sheet):
            # コンバート準備
            s = sheet
            self.row = self.yaml['table']['row'] - 1 # コンバート開始列
            self.setting_columns = {}
            # 存在するカラム処理
            yaml_column_set = set(self.columns)
            xls_column_set = set()
            for col in range(s.ncols):
                column = s.cell(self.row, col).value
                if column in self.columns:
                    #print column
                    xls_column_set.add(column)
                    default = self.columns[column]['default'] if 'default' in self.columns[column] else None
                    self.setting_columns[col] = SettingColumn(
                        name=self.columns[column]['name'],
                        type=self.columns[column]['type'],
                        default=default,
                    )
            # 存在しないカラム
            none_exist_columns = list(yaml_column_set - xls_column_set)
            self.settings_none_exist_columns = []
            for none_exist_column in none_exist_columns:
                column = none_exist_column
                try:
                    default = self.columns[column]['default']
                    self.settings_none_exist_columns.append(
                        SettingColumn(
                            name=self.columns[column]['name'],
                            type=self.columns[column]['type'],
                            default=default,
                            )
                        )
                except KeyError:
                    print u'存在しないカラム[%s]を指定されているが、デフォルト値が存在しない' % (column)
                    raise
            self.none_exist_columns = none_exist_columns
            
            # 拡張外部データ
            import_dict = None
            if 'import' in self.yaml['table']:
                import_file = self.path + self.yaml['table']['import']
                try:
                    f = open(import_file)
                    y = yaml.load(f.read())
                    import_dict = y
                except IOError:
                    print >>sys.stderr, u'[%s]を開くことができない' % (import_file)
            else:
                pass
            self.import_dict = import_dict
    


    def get_setting_column(self, row, column):
        if row <= self.row:
            return None
        if column in self.setting_columns:
            return self.setting_columns[column]
        return None

def xls2fix(s, settings, output_filename):
    fixture_list = []
    # 与えられたyamlの設定をしておく
    for row in range(s.nrows):
        rows = []
        for col in range(s.ncols):
            rows.append(s.cell(row, col).value)
        if row <= settings.row:
            continue
        fields = {}
        id = 0
        for column, col in enumerate(rows):
            # Excelのカラムがコンバート対象かチェックする
            setting_column = settings.get_setting_column(row, column)
            if setting_column:
                # コンバート対象カラム
                value = col
                if setting_column.type == 'datetime':
                    if value != '':
                        value = str(datetime.datetime(*xldate_as_tuple(value, 0)))
                    else:
                        value = None
                elif setting_column.type == 'char':
                    pass
                elif setting_column.type == 'int':
                    try:
                        if col == u'':
                            value = 0
                        else:
                            value = int(col)
                    except ValueError, UnicodeEncodeError:
                        # 置換できなかった場合、import_dictの中に変換可能なカラムがあるかをチェック
                        if setting_column.name in settings.import_dict:
                            column_dict = settings.import_dict[setting_column.name]
                            try:
                                value = column_dict[col]
                            except KeyError:
                                print >>sys.stderr, u'%s:%sはintでなくdictを使っても変換できない' % (cellnameabs(row, column), col)
                        else:
                            print >>sys.stderr, u'%s:%sはintに変換できない' % (cellnameabs(column, row), col)
                            raise
                            value = 0
                elif setting_column.type == 'float':
                    try:
                        value = float(col)
                    except ValueError:
                        value = 0.0
                elif setting_column.type == 'foreign_key':
                    try:
                        value = int(col)
                    except ValueError:
                        value = None
                    if value == 0:
                        value = None
                elif setting_column.type == 'boolean':
                    if len(unicode(col)) == 0:
                        value = False
                    else:
                        value = True
                else:
                    print u'存在しないカラムタイプ[%s]を指定されている' % (setting_column.type)
                    raise

                if setting_column.name == 'id':
                    id = int(value)
                else:
                    fields[setting_column.name] = value

        # 未設定カラムを順番に処理する
        for setting_column in settings.settings_none_exist_columns:

            fields[setting_column.name] = str(setting_column.default)

        fixture_list.append({
                'model': settings.model,
                'pk': id,
                'fields': fields,
                })


    fp = open(output_filename, 'w')

    if True:
        fp.write( json.dumps(fixture_list, encoding='utf-8', indent=4 * ' ') )
    else:
        fp.write( yaml.dump(fixture_list, encoding='utf-8', allow_unicode=True) )
    fp.close()

def main():
    
    parser = ArgumentParser()
    parser.add_argument('input_filename')
    parser.add_argument('-y', '--yaml', dest='yaml_filename', help='yaml filename')
    parser.add_argument('-o', '--output', dest='output_filename', help='output filename')

    args = parser.parse_args()
    
    # 与えられたyamlを解析する
    if args.yaml_filename:
        try:
            f = open(args.yaml_filename)
            settings = SettingsYaml(yaml.load(f.read()))
            settings.path = path.dirname( path.abspath( args.yaml_filename ) ) + u'/'
        except IOError:
            print u'%s が存在しません' % (args.yaml_filename)
            return
    else:
        print u'-yでyamlを指定してください'
        return

    input_filename = args.input_filename
    if args.output_filename:
        output_filename = args.output_filename
    else:
        root, ext = os.path.splitext(input_filename)
        if True:
            output_filename = root + '.json'
        else:
            output_filename = root + '.yaml'

    wb = open_workbook(input_filename)
    print u'Convert... %s' % (input_filename)
    for s in wb.sheets():
        if settings.is_convert_sheet(s):
            print u'Sheet:%s -> %s' % (s.name, output_filename)
            settings.setting_convert_sheet(s)
            xls2fix(s, settings, output_filename)

if __name__ == '__main__':
    try:
        main()
    except:
        print >>sys.stderr, u'エラーが発生しました'
        raise
