# _*_ coding: UTF-8 _*_

import xlrd
import xlwt


class ExcelReader:

    @staticmethod
    def read(file=None, sheet=None):
        workbook = xlrd.open_workbook(file)
        sheet = workbook.sheet_by_name(sheet)
        rows = [row for row in sheet.get_rows()]
        headers = rows[0]
        print('头部信息:{}'.format(headers))
        result = []
        for row in rows:
            if row == headers:
                continue
            tmp = {}
            for i in range(0, len(headers)):
                tmp[headers[i]] = row[i]
            result.append(tmp)
        return result


class ExcelWriter:

    @staticmethod
    def write(file=None, sheet=None, obj=None):
        if not obj or not isinstance(obj, list):
            raise Exception('不支持的写入对象')
        workbook = xlwt.Workbook()
        sheet = workbook.add_sheet(sheet)

        headers = obj[0].keys()
        for h_index in range(0, len(headers)):
            sheet.write(0, h_index, list(headers)[h_index])

        for c_index in range(0, len(obj)):
            for h_index in range(0, len(headers)):
                sheet.write(c_index, h_index, obj[c_index][list(headers)[h_index]])

        workbook.save(file)


if __name__ == '__main__':
    result_list = ExcelReader.read('../data/川信退款名单2.xls', 'Sheet3')
    print(result_list)
    ExcelWriter.write('../data/test.xls', 'Sheet1', result_list)
