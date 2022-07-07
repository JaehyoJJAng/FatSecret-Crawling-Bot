'''
사용 라이브러리 정보
Pywin32 : 엑셀 어플리케이션 다루기 위함
glob : 폴더 내의 엑셀 파일 경로 가져오기 위함
'''
# step1. 관련 모듈 및 패키지 import
import glob
import os.path
import win32com.client

class CollectionXLSX:
    @staticmethod
    def add_excel():
        # step2. win32com(Pywin32)를 이용해서 엑셀 어플리케이션 열기
        excel = win32com.client.Dispatch('Excel.Application')
        excel.Visible = False # 실제 작동하는 것을 보고 싶을 때 사용 (True , False)

        # step3. 엑셀 어플리케이션에 새로운 WorkBook 추가
        wb_new = excel.Workbooks.Add()

        # step4. glob 모듈로 원하는 폴더 내의 모든 xlsx 파일의 경로를 리스트로 반환
        list_filepath = glob.glob(r'C:\github\Web_Crawlling_Project\03_크몽프로젝트 객체지향 + 타입힌트 리팩토링\18_fast secret(패스트시크릿)\브랜드별식품정보\*.xlsx',recursive=True)

        # step5. 엑셀 시트를 추출하고 새로운 엑셀에 붙여넣는 반복문
        for idx,filepath in enumerate(list_filepath , 1):

            # 받아온 엑셀 파일의 경로를 이용해 엑셀 파일 열기
            wb = excel.Workbooks.Open(filepath)

            # 새로 만든 엑셀 파일에 추가
            # 추출할 wb.Worksheets('추출할 시트명').Copy(Before=붙여넣을 wb.Worksheets('기준 시트명')
            wb.Worksheets('레스토랑-체인점').Copy(Before=wb_new.Worksheets(f'취합본{idx}'))

        # step6. 취합한 엑셀 파일을 '통합 문서'라는 이름으로 저장
        savePath = os.path.abspath('FatSecret 취합 문서')
        fileName = '통합 문서.xlsx'
        wb_new.SaveAs(os.path.join(savePath,fileName))

        # step7. 켜져있는 엑셀 및 어플리케이션 모두 종료
        excel.Quit()


if __name__ == '__main__':
    app = CollectionXLSX()

    app.add_excel()
