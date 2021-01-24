import xlwt
import xlrd

file = 'China Lake_v1.xls'

#设置表格样式
def set_style(name,height,bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


if __name__ == '__main__':
    f = xlwt.Workbook()
    sheet2 = f.add_sheet('Lake_v2', cell_overwrite_ok=True)
    row0 = ["Year", "Month", "CHLA", "TEMPERATURE", "Total P"]
    # 写第一行
    for i in range(0, len(row0)):
        sheet2.write(0, i, row0[i], set_style('Times New Roman', 220, True))


    wb = xlrd.open_workbook(filename='Lake_v1.xlsx')  # 打开文件
    print(wb.sheet_names())  # 获取所有表格名字

    sheet1 = wb.sheet_by_index(3)  # 通过索引获取表格


    j=1

    year_ori = year = sheet1.cell_value(1, 5)
    month_ori = month = sheet1.cell_value(1, 6)
    day_ori = day = sheet1.cell_value(1, 7)

    #每月数据
    chla=tem=total=0
    amount_chla = amount_tem = amount_total = 0

    #每天数据
    sub_chla = sub_tem = sub_total = 0
    subAm_chla = subAm_tem = subAm_total = 0

    #如果存在某天同时包含三项度量数据时的数据
    exist_chla = exist_tem = exist_total = 0
    existAm = 0

    for i in range(1, sheet1.nrows):
        year = sheet1.cell_value(i, 5)
        day = sheet1.cell_value(i, 7)
        month = sheet1.cell_value(i, 6)

        if year_ori == year:

            if month_ori == month:

                if day_ori == day:

                    if sheet1.cell_value(i, 9) != '':
                        sub_chla = sub_chla + sheet1.cell_value(i, 9)
                        subAm_chla = subAm_chla + 1
                    elif sheet1.cell_value(i, 10) != '':
                        sub_tem = sub_tem + sheet1.cell_value(i, 10)
                        subAm_tem = subAm_tem + 1
                    else:
                        sub_total = sub_total + sheet1.cell_value(i, 11)
                        subAm_total = subAm_total + 1
                else:
                    if sub_chla != 0 and sub_tem != 0 and sub_total != 0:
                        exist_chla = exist_chla + sub_chla/subAm_chla
                        exist_tem = exist_tem + sub_tem / subAm_tem
                        exist_total = exist_total + sub_total / subAm_total
                        existAm = existAm + 1

                    chla = chla + sub_chla
                    tem = tem + sub_tem
                    total = total + sub_total
                    amount_chla = amount_chla + subAm_chla
                    amount_tem = amount_tem + subAm_tem
                    amount_total = amount_total + subAm_total

                    sub_chla = sub_tem = sub_total = 0
                    subAm_chla = subAm_tem = subAm_total = 0

                    if sheet1.cell_value(i, 9) != '':
                        sub_chla = sub_chla + sheet1.cell_value(i, 9)
                        subAm_chla = subAm_chla + 1
                    elif sheet1.cell_value(i, 10) != '':
                        sub_tem = sub_tem + sheet1.cell_value(i, 10)
                        subAm_tem = subAm_tem + 1
                    else:
                        sub_total = sub_total + sheet1.cell_value(i, 11)
                        subAm_total = subAm_total + 1



            else:
                if sub_chla != 0 and sub_tem != 0 and sub_total != 0:
                    exist_chla = exist_chla + sub_chla / subAm_chla
                    exist_tem = exist_tem + sub_tem / subAm_tem
                    exist_total = exist_total + sub_total / subAm_total
                    existAm = existAm + 1

                chla = chla + sub_chla
                tem = tem + sub_tem
                total = total + sub_total
                amount_chla = amount_chla + subAm_chla
                amount_tem = amount_tem + subAm_tem
                amount_total = amount_total + subAm_total

                sheet2.write(j, 0, year_ori)
                sheet2.write(j, 1, month_ori)
                if existAm == 0:
                    if amount_chla != 0:
                        sheet2.write(j, 2, chla/amount_chla)
                    if amount_tem != 0:
                        sheet2.write(j, 3, tem/amount_tem)
                    if amount_total != 0:
                        sheet2.write(j, 4, total/amount_total)
                else:
                    sheet2.write(j, 2, exist_chla / existAm)
                    sheet2.write(j, 3, exist_tem / existAm)
                    sheet2.write(j, 4, exist_total / existAm)

                j = j+1

                chla = tem = total = 0
                amount_chla = amount_tem = amount_total = 0

                sub_chla = sub_tem = sub_total = 0
                subAm_chla = subAm_tem = subAm_total = 0

                exist_chla = exist_tem = exist_total = 0
                existAm = 0

                if sheet1.cell_value(i, 9) != '':
                    sub_chla = sub_chla + sheet1.cell_value(i, 9)
                    subAm_chla = subAm_chla + 1
                elif sheet1.cell_value(i, 10) != '':
                    sub_tem = sub_tem + sheet1.cell_value(i, 10)
                    subAm_tem = subAm_tem + 1
                else:
                    sub_total = sub_total + sheet1.cell_value(i, 11)
                    subAm_total = subAm_total + 1


        else:
            if sub_chla != 0 and sub_tem != 0 and sub_total != 0:
                exist_chla = exist_chla + sub_chla / subAm_chla
                exist_tem = exist_tem + sub_tem / subAm_tem
                exist_total = exist_total + sub_total / subAm_total
                existAm = existAm + 1

            chla = chla + sub_chla
            tem = tem + sub_tem
            total = total + sub_total
            amount_chla = amount_chla + subAm_chla
            amount_tem = amount_tem + subAm_tem
            amount_total = amount_total + subAm_total

            sheet2.write(j, 0, year_ori)
            sheet2.write(j, 1, month_ori)
            if existAm == 0:
                if amount_chla != 0:
                    sheet2.write(j, 2, chla / amount_chla)
                if amount_tem != 0:
                    sheet2.write(j, 3, tem / amount_tem)
                if amount_total != 0:
                    sheet2.write(j, 4, total / amount_total)
            else:
                sheet2.write(j, 2, exist_chla / existAm)
                sheet2.write(j, 3, exist_tem / existAm)
                sheet2.write(j, 4, exist_total / existAm)

            j = j + 1

            chla = tem = total = 0
            amount_chla = amount_tem = amount_total = 0

            sub_chla = sub_tem = sub_total = 0
            subAm_chla = subAm_tem = subAm_total = 0

            exist_chla = exist_tem = exist_total = 0
            existAm = 0

            if sheet1.cell_value(i, 9) != '':
                sub_chla = sub_chla + sheet1.cell_value(i, 9)
                subAm_chla = subAm_chla + 1
            elif sheet1.cell_value(i, 10) != '':
                sub_tem = sub_tem + sheet1.cell_value(i, 10)
                subAm_tem = subAm_tem + 1
            else:
                sub_total = sub_total + sheet1.cell_value(i, 11)
                subAm_total = subAm_total + 1

        year_ori = year
        month_ori = month
        day_ori = day

    if sub_chla != 0 and sub_tem != 0 and sub_total != 0:
        exist_chla = exist_chla + sub_chla / subAm_chla
        exist_tem = exist_tem + sub_tem / subAm_tem
        exist_total = exist_total + sub_total / subAm_total
        existAm = existAm + 1

    chla = chla + sub_chla
    tem = tem + sub_tem
    total = total + sub_total
    amount_chla = amount_chla + subAm_chla
    amount_tem = amount_tem + subAm_tem
    amount_total = amount_total + subAm_total

    sheet2.write(j, 0, year_ori)
    sheet2.write(j, 1, month_ori)
    if existAm == 0:
        if amount_chla != 0:
            sheet2.write(j, 2, chla / amount_chla)
        if amount_tem != 0:
            sheet2.write(j, 3, tem / amount_tem)
        if amount_total != 0:
            sheet2.write(j, 4, total / amount_total)
    else:
        sheet2.write(j, 2, exist_chla / existAm)
        sheet2.write(j, 3, exist_tem / existAm)
        sheet2.write(j, 4, exist_total / existAm)

    f.save('Lake_v2.xls')