import streamlit as st
import pandas as pd
import openpyxl
import io

# 1. 웹페이지 기본 설정
st.set_page_config(page_title="AI스캐너 정산 자동화 시스템", page_icon="🤖", layout="centered")

st.title("🤖 AI스캐너 월별 정산 자동화 시스템")
st.markdown("당월 정산 데이터와 마스터 정보 엑셀 2개만 올려주시면, 양식 템플릿에 맞게 정산 내역서가 자동으로 채워집니다.")

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

# 4. 파일 업로드 섹션 (필요 엑셀 파일 2개 + 양식 템플릿)
st.subheader("2️⃣ 엑셀 파일 업로드")

col1, col2 = st.columns(2)

with col1:
    file_cur = st.file_uploader("1) 당월 정산 데이터 (`260728_AI스캐너_정산_...`)", type=["xlsx", "xls"])
    file_master = st.file_uploader("2) 마스터 정산 정보 (`▸AI스캐너_정산정보.xlsx`)", type=["xlsx", "xls"])

with col2:
    file_tpl = st.file_uploader("3) 정산 양식 템플릿 (`▸(smpl)AI스캐너_정산내역.xlsx`)", type=["xlsx", "xls"])

st.divider()

# 5. 정산 처리 및 데이터 자동 채우기
st.subheader("3️⃣ 정산 실행 및 결과 다운로드")

if st.button("🚀 정산 내역서 자동 생성 시작", type="primary", use_container_width=True):
    if not file_cur or not file_master or not file_tpl:
        st.warning("⚠️ 엑셀 파일 3개를 모두 업로드한 후 실행해 주세요.")
    else:
        with st.spinner("⏳ 데이터를 파싱하고 템플릿 서식에 데이터를 채우는 중입니다..."):
            try:
                # 1) 템플릿 파일 로드
                wb = openpyxl.load_workbook(file_tpl)
                ws = wb[wb.sheetnames[0]]
                
                # 시트명 및 헤더 연월 업데이트
                ws.title = m3
                ws['C2'] = f"{m3} 대리점 정산(VAT포함)"
                ws['J5'] = f"{m1}(VAT포함)"
                ws['K5'] = f"{m2}(VAT포함)"
                ws['L5'] = f"{m3}(VAT포함)"
                ws['N5'] = f"{m3}(VAT포함)"
                
                # 두 번째 시트(Sheet1)가 있으면 VLOOKUP 수식의 시트명도 변경
                if len(wb.sheetnames) > 1:
                    ws2 = wb[wb.sheetnames[1]]
                    ws2['C2'] = f"{m3} 대리점 정산(VAT포함)"
                    for r_idx in range(3, 30):
                        ws2[f'C{r_idx}'] = f"=VLOOKUP(B:B,'{m3}'!B:N,12,0)"

                # 2) 마스터 정산정보 파싱 (CMS 시트)
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

                # 3) 당월 정산 데이터 파싱
                df_cur = pd.read_excel(file_cur, sheet_name=0)
                
                store_records = []
                unmapped_stores = []

                # 당월 데이터 행순회 (Row 4부터 데이터 시작)
                for i in range(4, len(df_cur)):
                    row = df_cur.iloc[i]
                    store_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
                    if not store_name or store_name == '합계' or store_name == 'nan':
                        continue

                    # 주요 데이터 추출 (VAT별도 -> VAT포함 변환)
                    qty = row.iloc[15] if pd.notna(row.iloc[15]) else 1
                    cost_ex = row.iloc[18] if pd.notna(row.iloc[18]) else 0
                    pb_ex = row.iloc[19] if pd.notna(row.iloc[19]) else 0
                    
                    cost_inc = round(float(cost_ex) * 1.1)
                    pb_inc = round(float(pb_ex) * 1.1)

                    # 대리점 및 마스터 정보 결합
                    master_info = master_map.get(store_name, {})
                    agency = str(row.iloc[11]).strip() if pd.notna(row.iloc[11]) else master_info.get('agency', '미매핑 대리점')
                    biz_no = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else master_info.get('biz_no', '')
                    sales_rep = str(row.iloc[14]).strip() if pd.notna(row.iloc[14]) else master_info.get('sales_rep', '')

                    if not master_info and agency == '미매핑 대리점':
                        unmapped_stores.append(store_name)

                    store_records.append({
                        'agency': agency,
                        'store_name': store_name,
                        'biz_no': biz_no,
                        'sales_rep': sales_rep,
                        'qty': qty,
                        'cost_inc': cost_inc,
                        'pb_inc': pb_inc
                    })

                # 대리점별 그룹화 정렬
                df_stores = pd.DataFrame(store_records)
                df_stores.sort_values(by=['agency', 'store_name'], inplace=True)

                # 4) 빈 템플릿(Row 7부터)에 데이터 채워넣기
                start_row = 7
                agency_start_row_map = {}

                for idx, item in enumerate(df_stores.to_dict('records')):
                    r = start_row + idx
                    agency = item['agency']

                    # 대리점 이름은 대리점의 첫 번째 매장 행에만 표시
                    if agency not in agency_start_row_map:
                        ws[f'B{r}'] = agency
                        agency_start_row_map[agency] = [r]
                    else:
                        agency_start_row_map[agency].append(r)

                    ws[f'C{r}'] = item['store_name']
                    ws[f'D{r}'] = item['biz_no']
                    ws[f'E{r}'] = item['sales_rep']
                    ws[f'F{r}'] = item['qty']
                    ws[f'G{r}'] = item['cost_inc']
                    ws[f'H{r}'] = item['pb_inc']
                    ws[f'J{r}'] = item['cost_inc']  # 이전월1 CMS
                    ws[f'K{r}'] = item['cost_inc']  # 이전월2 CMS
                    ws[f'L{r}'] = item['cost_inc']  # 당월 CMS
                    ws[f'M{r}'] = item['pb_inc']    # 당월 정산금

                # 대리점별 최종 정산 합계금액(N열)을 첫 번째 행에 작성
                for agency, rows in agency_start_row_map.items():
                    first_r = rows[0]
                    # 해당 대리점에 속한 모든 매장의 정산금(M열) 합계
                    total_agency_pb = sum(df_stores.iloc[r - start_row]['pb_inc'] for r in rows)
                    ws[f'N{first_r}'] = total_agency_pb

                # 5) 메모리 버퍼에 저장 후 다운로드 제공
                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)
                
                st.success(f"🎉 `{m3}` 정산 내역서 데이터 채우기가 완료되었습니다!")
                
                if unmapped_stores:
                    st.warning(f"⚠️ 당월 데이터 미매핑 매장 ({len(unmapped_stores)}건): {', '.join(unmapped_stores)}")

                st.download_button(
                    label=f"📥 AI스캐너_정산내역_{m3}.xlsx 다운로드",
                    data=output_buffer,
                    file_name=f"AI스캐너_정산내역_{m3}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            except Exception as e:
                st.error(f"❌ 정산 작업 중 오류가 발생했습니다: {e}")
