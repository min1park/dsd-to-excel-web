import streamlit as st
import tempfile
import os
from dsd_to_excel import convert_dsd_to_excel

# 페이지 기본 설정
st.set_page_config(
    page_title="DSD to Excel 변환기",
    page_icon="📊",
    layout="centered"
)

# 메인 타이틀 및 설명
st.title("📊 DART DSD → Excel 변환기")
st.markdown("""
DART 전자공시 감사보고서(**.dsd**) 파일을 업로드하시면 
깔끔하게 정리된 **Excel(.xlsx)** 파일로 변환해 드립니다.
""")

st.divider()

# 파일 업로드 컴포넌트
uploaded_file = st.file_uploader("DSD 파일을 이곳에 드래그하거나 선택하세요", type=["dsd"])

if uploaded_file is not None:
    # 원본 파일명에서 확장자 분리 (예: report.dsd -> report)
    original_filename = os.path.splitext(uploaded_file.name)[0]
    
    st.info(f"📁 '{uploaded_file.name}' 파일이 업로드되었습니다. 변환 버튼을 눌러주세요.")
    
    # 변환 실행 버튼
    if st.button("🚀 Excel로 변환하기", use_container_width=True):
        
        # 처리 중 스피너 표시
        with st.spinner('파일을 분석하고 변환하는 중입니다. 잠시만 기다려주세요...'):
            
            # 임시 파일 경로를 담을 변수 초기화
            tmp_dsd_path = ""
            tmp_xlsx_path = ""
            
            try:
                # 1. 업로드된 데이터를 임시 DSD 파일로 저장
                with tempfile.NamedTemporaryFile(delete=False, suffix='.dsd') as tmp_dsd:
                    tmp_dsd.write(uploaded_file.getvalue())
                    tmp_dsd_path = tmp_dsd.name
                
                # 2. 결과물(Excel)이 저장될 임시 경로 지정
                tmp_xlsx_path = tmp_dsd_path.replace('.dsd', '.xlsx')
                
                # 3. 기존 로직(convert_dsd_to_excel) 실행
                convert_dsd_to_excel(tmp_dsd_path, tmp_xlsx_path)
                
                # 4. 변환 성공 시 파일 읽기 및 다운로드 버튼 제공
                with open(tmp_xlsx_path, "rb") as excel_file:
                    excel_data = excel_file.read()
                    
                st.success("🎉 성공적으로 변환되었습니다! 아래 버튼을 눌러 다운로드하세요.")
                
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=excel_data,
                    file_name=f"{original_filename}_변환완료.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True
                )
                
            except Exception as e:
                # 오류 발생 시 사용자에게 알림
                st.error(f"⚠️ 변환 중 오류가 발생했습니다: {str(e)}")
                
            finally:
                # 5. 작업 완료 후 서버의 임시 파일들 삭제 (용량 관리)
                if tmp_dsd_path and os.path.exists(tmp_dsd_path):
                    os.remove(tmp_dsd_path)
                if tmp_xlsx_path and os.path.exists(tmp_xlsx_path):
                    os.remove(tmp_xlsx_path)