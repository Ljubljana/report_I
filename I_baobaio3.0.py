import pandas as pd
import re
# import xlrd

excel_h = pd.ExcelFile('./h.xlsx' )
h0 = pd.read_excel('./0.xls')

for w in range(0, len(excel_h.sheet_names)):
    G = globals()
    G[str(w)] = pd.read_excel('./h.xlsx', sheet_name=excel_h.sheet_names[w])

    for i in h0.index:
        deser = h0['"descr"'].at[i]
        zdz = h0['"zdz"'].at[i]
        hz = re.findall(r'([\u2E80-\u9FFF]+)', deser)   #匹配所有汉字
        zhanming_0 = hz[0]  #列表第一个是站名,如白塔变
        fg = re.split(r'[.]', deser)    #提取如：912母联开关A相电流
        dhtq = re.findall(r'[a-zA-Z0-9]+', fg[-1])   #'+'表示字符串而不是单个字符
        diaohao_0 = dhtq[0]    #如：924y

        for q in G[str(w)].index:
            diaohao_h = re.findall(r'[a-zA-Z0-9]+', str(G[str(w)].iloc[q, 2]))
            if zhanming_0 in excel_h.sheet_names[w] and diaohao_0 == diaohao_h[0]:
                G[str(w)].iloc[q, 5] = h0.iloc[i, 2]        #[q, 4]是同期运行最大电流，[q, 5]是当前线路最大电流
                if G[str(w)].iloc[q, 5] >= 400:     #这两行判断电流是否大于400A
                    print(str(G[str(w)].iloc[q, 1]) + str(G[str(w)].iloc[q, 2]) + str(G[str(w)].iloc[q, 5]))


writer = pd.ExcelWriter('./out.xlsx')
for e in range(0, len(excel_h.sheet_names)):
    G[str(e)].to_excel(writer, sheet_name=excel_h.sheet_names[e], index=False)
writer.save()
writer.close()
