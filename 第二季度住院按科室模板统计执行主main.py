import pandas as pd
import method3 as t

if __name__ == '__main__':
    # 自定义路径
    excel_path = r"excels\规则.xlsx"
    excel_path1 = r"excels\第二季度住院病人来源.xlsx"
    sheets = ['CDS-成都市','SCS-四川省','外省']
    sheet1 = u'Sheet1'
    COL = 1
    # 获取数据源数据
    df = pd.read_excel(excel_path1, sheet_name=sheet1)
    arr = t.Get_dict()
    # print(arr)
    writer = pd.ExcelWriter("%s统计结果科室处理.xlsx" % excel_path1)
    df_1 = df[df['省']==sheets[1]]
    df_2 = df_1[df_1['市']==sheets[0]]

    result0= pd.DataFrame()
    for arr_ in arr[0].values():
        pv1 = df_2[df_2['地']==arr_]
        result0 = pv1.append(result0)
    pv_result = pd.pivot_table(result0,index=['出院科室','省','市','地','医保类型'],values='出院人数',aggfunc='sum')

    pv_result.to_excel(writer, sheet_name=sheet1, startrow=2, startcol=COL)
    COL = COL+7
    result1 = pd.DataFrame()

    for arr_ in arr[1].values():
        pv1 = df_1[df_1['市']==arr_]
        result1 = pv1.append(result1)
    pv_result1 = pd.pivot_table(result1, index=['出院科室', '省','市','医保类型'], values='出院人数',aggfunc='sum')
    pv_result1.to_excel(writer, sheet_name=sheet1, startrow=2, startcol=COL)
    COL = COL + 6
    result2 = pd.DataFrame()
    for arr_ in arr[2].values():
        pv1 = df[df['省']==arr_]
        result2 = pv1.append(result2)
    pv_result2 = pd.pivot_table(result2, index=['出院科室','省','医保类型'], values='出院人数',aggfunc='sum')
    pv_result2.to_excel(writer, sheet_name=sheet1, startrow=2, startcol=COL)
    writer.save()

