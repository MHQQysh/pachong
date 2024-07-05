import time

import openpyxl
import pandas as pd
from DrissionPage._pages.web_page import WebPage

page = WebPage()
data_list = []

# excel保存
# 1创建文档
wb = openpyxl.Workbook()
# 2创建表, 一张Excel表里有几张表
#sheet1 = wb.active
sheet1 = wb.create_sheet('Sheet1')  #一张表里写多个表
sheet1.append(['Journal','ISSN','eISSN','Category','Edition','Total Citation','2023 JIF','JIF Quartile','2023 JCI','% of OA Gold'])
url = 'https://jcr.clarivate.com/jcr/browse-journals?app=jcr&referrer=target%3Dhttps:%2F%2Fjcr.clarivate.com%2Fjcr%2Fbrowse-journals&Init=Yes&authCode=IpKG2TDzzWo0HLJwg0P1Bsez3Htk5-gJLHhtaUjE7ag&SrcApp=IC2LS'

page.get(url)
time.sleep(0)
# print(page.html)


with open('shuju.csv','w',encoding='utf-8-sig')as f:
    f.write('Journal,ISSN,eISSN,Category,Edition,Total Citation,2023 JIF,JIF Quartile,2023 JCI,% of OA Gold\n')

    for i in range(1,1000):   #页数1-10页
        page.ele('x://button[@aria-label="Next page"]').click(by_js=True)
        page.ele('x://button[@aria-label="Next page"]').click(by_js=True)

        sj_list= page.eles('x://mat-table[@role="table"]/mat-row')
        print(sj_list)
        for sj in sj_list:
            #  Journal name
            ju = sj.ele('x:.//span[@class="table-cell-journalName ng-star-inserted"]').text
            #print(ju)
            # ISSN
            ISSN = sj.ele('x:.//span[@class="table-cell-issn NAN ng-star-inserted"]').text

            #  eISSN
            eISSN = sj.ele('x:.//span[@class="table-cell-eissn NAN ng-star-inserted"]').text


            # Category
            if sj.ele('x:.//span[@class="multiple table-cell-category ng-star-inserted"]/@title'):
                Category = ''.join(sj.ele('x:.//span[@class="multiple table-cell-category ng-star-inserted"]/@title')).replace(',','，').strip()
            else:
                Category = ''.join(sj.ele('x:.//span[@class="table-cell-category ng-star-inserted"]').text).replace(',','，').strip()
            #  Edition
            Edition = ''.join(sj.ele('x:.//span[@class="table-cell-edition ng-star-inserted"]').text).replace(',','，').strip()

            # Total Citation
            Total = ''.join(sj.ele('x:.//span[@class="table-cell-totalCites ng-star-inserted"]').text).replace(',','，').strip()

            # 2023 JIF
            JIF_2013 = sj.ele('x:.//span[@class="table-cell-jif2019 ng-star-inserted"]/@title')

            # JIF Quartile
            # JIF_Quartile = sj.ele('x:.//span[@class="table-cell-quartile ng-star-inserted"]/@title')
            JIF_Quartile = 'Q1'
            # 2023 JCI
            JCI_2023 = sj.ele('x:.//span[@class="table-cell-jci ng-star-inserted"]/@title')

            # % of OA Gold
            ofOAGold = sj.ele('x:.//span[@class="table-cell-percentageOAGold ng-star-inserted"]/@title')

            f.write(f'{ju},{ISSN},{eISSN},{Category},{Edition},{Total},{JIF_2013},{JIF_Quartile},{JCI_2023},{ofOAGold}\n')
            sheet1.append([ju,ISSN,eISSN,Category,Edition,Total,JIF_2013,JIF_Quartile,JCI_2023,ofOAGold])

            wb.save('ss.xlsx')
            print(f'第{i}页{ju}写入成功')
        # 点击下一页
        page.ele('x://button[@aria-label="Next page"]').click(by_js=True)







        # data_list.append({
        # Journal,ISSN,eISSN,Category,Edition,Total Citation,2023 JIF,JIF Quartile,2023 JCI,% of OA Gold
        #     'Journal name': ju,
        #     'ISSN': ISSN,
        #     'eISSN': eISSN,
        #     'Category': Category,
        #     'Edition': Edition,
        #     'Total Citation': Total,
        #     '2023 JIF': JIF_2013,
        #     'JIF Quartile': JIF_Quartile,
        #     '2023 JCI': JCI_2023,
        #     '% of OA Gold': ofOAGold,
        #
        #
        # })
        # print(data_list)
        # # 每行的单位列自动添加
        # df2 = pd.DataFrame(data=data_list)
        # # print(df)
        #
        # df2.to_excel(f'数据.xlsx', index=False)
        #
        #print(f'写入完毕')

