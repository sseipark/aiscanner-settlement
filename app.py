import streamlit as st
import pandas as pd
import openpyxl
import io

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="AI스캐너 정산 자동화 시스템", page_icon="🤖", layout="centered")

st.title("🤖 AI스캐너 월별 정산 자동화 시스템")
st.markdown("매월 정산 데이터와 마스터 정보 엑셀을 업로드하면, 서식이 보존된 정산 내역서를 자동으로 생성해 줍니다.")

st.divider()

# 2. 최근 3개월 연월 자동 계산 함수
def get_recent_3months(yymm_str):
    yy = int("20" + yymm_str[:2])
    mm = int(yymm_str[2:])
    
    months = []
    for i in range(2, -1, -1):
        m = mm - i
        y = yy
        if m <= 0:
            m += 12
            y -= 1
        months.append(f"{str(y)[2:]}{m:02d}")
    return months

# 3. 정산 대상 월 입력 섹션
st.subheader("1️⃣ 정산 대상 월 입력")
target_yymm = st.text_input("정산 연월 4자리를 입력하세요 (예: 7월 정산 -> 2607, 8월 정산 -> 2608)", value="2607")

if target_yymm and len(target_yymm) == 4 and target_yymm.isdigit():
    m1, m2, m3 = get_recent_3months(target_yymm)
    st.info(f"📅 **적용될 정산 월:** `{m3}` | **최근 3개월 헤더:** `{m1}`, `{m2}`, `{m3}`")
else:
    st.error("⚠️ 정산 연월을 숫자 4자리(예: 2607)로 바르게 입력해 주세요.")

st.divider()

# 4. 파일 업로드 섹션
st.subheader("2️⃣ 필수 엑셀 파일 3개 업로드")

col1, col2 = st.columns(2)

with col1:
    file_cur = st.file_uploader("1) 당월 정산 데이터 (`260728_AI스캐너_정산_...`)", type=["xlsx", "xls"])
    file_master = st.file_uploader("2) 마스터 정산 정보 (`▸AI스캐너_정산정보.xlsx`)", type=["xlsx", "xls"])

with col2:
    file_tpl = st.file_uploader("3) 서식 템플릿 (`AI스캐너_정산내역_2606_수정.xlsx`)", type=["xlsx", "xls"])

st.divider()

# 5. 정산 처리 및 다운로드 섹션
st.subheader("3️⃣ 정산 실행 및 결과 다운로드")

if st.button("🚀 정산 내역서 자동 생성 시작", type="primary", use_container_width=True):
    if not file_cur or not file_master or not file_tpl:
        st.warning("⚠️ 엑셀 파일 3개를 모두 업로드한 후 실행해 주세요.")
    else:
        with st.spinner("⏳ 데이터를 분석하고 엑셀 서식을 업데이트하는 중입니다..."):
            try:
                # 1) 템플릿 로드 및 시트명/헤더 월 변경
                wb = openpyxl.load_workbook(file_tpl)
                ws = wb[wb.sheetnames[0]]
                
                ws.title = m3
                ws['C1'] = f"{m3} 대리점 정산(VAT포함)"
                ws['J4'] = f"{m1}(VAT포함)"
                ws['K4'] = f"{m2}(VAT포함)"
                ws['L4'] = f"{m3}(VAT포함)"
                ws['N4'] = f"{m3}(VAT포함)"
                
                # 2) 데이터 파싱 및 마스터 정보 매핑
                df_cur = pd.read_excel(file_cur, sheet_name=0)
                df_master_cms = pd.read_excel(file_master, sheet_name='CMS')
                
                master_map = {}
                for idx, row in df_master_cms.iterrows():
                    store_name = str(row['Unnamed: 1']).strip() if pd.notna(row['Unnamed: 1']) else ''
                    if store_name and store_name != '엑셀':
                        master_map[store_name] = {
                            'agency': str(row['Unnamed: 6']).strip() if pd.notna(row['Unnamed: 6']) else '',
                            'biz_no': str(row['Unnamed: 5']).strip() if pd.notna(row['Unnamed: 5']) else '',
                            'sales_rep': str(row['Unnamed: 8']).strip() if pd.notna(row['Unnamed: 8']) else ''
                        }

                # 3) 7월 정산 수량 및 매장 정보 채우기
                row_idx = 6
                unmapped_stores = []
                
                while True:
                    agency_cell = ws[f'B{row_idx}'].value
                    store_cell = ws[f'C{row_idx}'].value
                    
                    if agency_cell is None and store_cell is None:
                        break
                        
                    store_name = str(store_cell).strip() if store_cell else ''
                    if store_name:
                        matched_cur = df_cur[df_cur.iloc[:, 2].astype(str).str.strip() == store_name]
                        if not matched_cur.empty:
                            cur_row = matched_cur.iloc[0]
                            qty = cur_row.iloc[15] if pd.notna(cur_row.iloc[15]) else 1
                            
                            if store_name in master_map:
                                info = master_map[store_name]
                                if info['biz_no']: ws[f'D{row_idx}'] = info['biz_no']
                                if info['sales_rep']: ws[f'E{row_idx}'] = info['sales_rep']
                            
                            ws[f'F{row_idx}'] = qty
                        else:
                            unmapped_stores.append(store_name)
                            
                    row_idx += 1

                # 4) 메모리에 파일 저장
                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)
                
                st.success(f"🎉 `{m3}` 정산 내역 생성이 완료되었습니다!")
                
                if unmapped_stores:
                    st.warning(f"⚠️ 당월 데이터 미매핑 매장 ({len(unmapped_stores)}건): {', '.join(unmapped_stores)}")

                # 5) 다운로드 버튼
                st.download_button(
                    label=f"📥 AI스캐너_정산내역_{m3}.xlsx 다운로드",
                    data=output_buffer,
                    file_name=f"AI스캐너_정산내역_{m3}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ 정산 작업 중 오류가 발생했습니다: {e}")
