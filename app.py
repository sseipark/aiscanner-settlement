import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment, PatternFill
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

# 안전한 숫자 변환 함수
def safe_float(val, default=0.0):
    try:
        if pd.isna(val): return default
        return float(val)
    except (ValueError, TypeError):
        return default

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
st.subheader("2️⃣ 엑셀 파일 업로드")

col1, col2 = st.columns(2)

with col1:
    file_cur = st.file_uploader("1) 당월 정산 데이터 (`260728_AI스캐너_정산_...`)", type=["xlsx", "xls"])
    file_master = st.file_uploader("2) 마스터 정산 정보 (`▸AI스캐너_정산정보.xlsx`)", type=["xlsx", "xls"])

with col2:
    file_tpl = st.file_uploader("3) 정산 양식 템플릿 (`▸(smpl)AI스캐너_정산내역.xlsx`)", type=["xlsx", "xls"])

st.divider()

# 5. 정산 처리 및 양식 자동 확장 데이터 채우기
st.subheader("3️⃣ 정산 실행 및 결과 다운로드")

if st.button("🚀 정산 내역서 자동 생성 시작", type="primary", use_container_width=True):
    if not file_cur or not file_master or not file_tpl:
        st.warning("⚠️ 엑셀 파일 3개를 모두 업로드한 후 실행해 주세요.")
    else:
        with st.spinner("⏳ 필터링 및 엑셀 서식을 맞추어 정산서를 생성 중입니다..."):
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
                
                if len(wb.sheetnames) > 1:
                    ws2 = wb[wb.sheetnames[1]]
                    ws2['C2'] = f"{m3} 대리점 정산(VAT포함)"
                    for r_idx in range(3, 35):
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

                # 3) 당월 정산 데이터 파싱 (운영 상태가 '청구' 및 '미청구'인 매장만 선택)
                df_cur = pd.read_excel(file_cur, sheet_name=0)
                
                store_records = []
                unmapped_stores = []

                for i in range(4, len(df_cur)):
                    row = df_cur.iloc[i]
                    store_name = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
                    status = str(row.iloc[16]).strip() if pd.notna(row.iloc[16]) else ''

                    # ✨ 핵심 1: '청구' 및 '미청구' 대상 매장만 정산에 포함 (계약해지 매장 제외)
                    if not store_name or store_name == '합계' or store_name == 'nan':
                        continue
                    if status not in ['청구', '미청구']:
                        continue

                    qty_val = row.iloc[15]
                    qty = int(safe_float(qty_val, 1)) if safe_float(qty_val, 1) > 0 else 1
                    
                    cost_ex = safe_float(row.iloc[18], 0.0)
                    pb_ex = safe_float(row.iloc[19], 0.0)
                    
                    cost_inc = round(cost_ex * 1.1)
                    pb_inc = round(pb_ex * 1.1)

                    note = ''
                    if isinstance(row.iloc[18], str) and not str(row.iloc[18]).replace('.','',1).isdigit():
                        note = str(row.iloc[18]).strip()
                    elif isinstance(row.iloc[19], str) and not str(row.iloc[19]).replace('.','',1).isdigit():
                        note = str(row.iloc[19]).strip()

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
                        'pb_inc': pb_inc,
                        'note': note
                    })

                df_stores = pd.DataFrame(store_records)
                df_stores.sort_values(by=['agency', 'store_name'], inplace=True)

                # 4) 데이터 채우기 및 서식/행 높이 보정
                start_row = 7
                records = df_stores.to_dict('records')
                total_records = len(records)
                end_data_row = start_row + total_records - 1

                thin_border = Border(
                    left=Side(style='thin', color='B0C4DE'),
                    right=Side(style='thin', color='B0C4DE'),
                    top=Side(style='thin', color='B0C4DE'),
                    bottom=Side(style='thin', color='B0C4DE')
                )

                agency_summary_map = {}

                for idx, item in enumerate(records):
                    r = start_row + idx
                    agency = item['agency']

                    # ✨ 핵심 2: 칸 높이 고정 (22pt)
                    ws.row_dimensions[r].height = 22

                    if agency not in agency_summary_map:
                        agency_summary_map[agency] = {'first_row': r, 'total_pb': 0}
                    agency_summary_map[agency]['total_pb'] += item['pb_inc']

                    ws[f'B{r}'] = agency
                    ws[f'C{r}'] = item['store_name']
                    ws[f'D{r}'] = item['biz_no']
                    ws[f'E{r}'] = item['sales_rep']
                    ws[f'F{r}'] = item['qty']
                    ws[f'G{r}'] = item['cost_inc']
                    ws[f'H{r}'] = item['pb_inc']
                    if item['note']:
                        ws[f'I{r}'] = item['note']
                    ws[f'J{r}'] = item['cost_inc']
                    ws[f'K{r}'] = item['cost_inc']
                    ws[f'L{r}'] = item['cost_inc']
                    ws[f'M{r}'] = item['pb_inc']

                    # 셀 스타일 정렬 및 테두리 보정
                    for col_idx in range(2, 15):
                        cell = ws.cell(row=r, column=col_idx)
                        cell.border = thin_border
                        cell.font = Font(name='맑은 고딕', size=10)
                        if col_idx in [2, 3, 4, 5, 9]:
                            cell.alignment = Alignment(horizontal='center', vertical='center')
                        else:
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                        if col_idx in [6, 7, 8, 10, 11, 12, 13, 14]:
                            cell.number_format = '#,##0'

                # 대리점 정산 합계(N열)는 대리점의 첫 번째 매장 행에 작성
                for agency, info in agency_summary_map.items():
                    ws[f'N{info["first_row"]}'] = info['total_pb']

                # 5) 맨 아래 합계(SUM) 행 생성
                sum_row = end_data_row + 1
                ws.row_dimensions[sum_row].height = 24
                
                ws[f'B{sum_row}'] = "합계"
                ws[f'F{sum_row}'] = f"=SUM(F7:F{end_data_row})"
                ws[f'G{sum_row}'] = f"=SUM(G7:G{end_data_row})"
                ws[f'H{sum_row}'] = f"=SUM(H7:H{end_data_row})"
                ws[f'J{sum_row}'] = f"=SUM(J7:J{end_data_row})"
                ws[f'K{sum_row}'] = f"=SUM(K7:K{end_data_row})"
                ws[f'L{sum_row}'] = f"=SUM(L7:L{end_data_row})"
                ws[f'M{sum_row}'] = f"=SUM(M7:M{end_data_row})"
                ws[f'N{sum_row}'] = f"=SUM(N7:N{end_data_row})"

                # 합계 행 강조 스타일
                sum_fill = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
                for col_idx in range(2, 15):
                    cell = ws.cell(row=sum_row, column=col_idx)
                    cell.border = thin_border
                    cell.fill = sum_fill
                    cell.font = Font(name='맑은 고딕', size=10, bold=True)
                    if col_idx in [6, 7, 8, 10, 11, 12, 13, 14]:
                        cell.number_format = '#,##0'

                # 6) 메모리 버퍼 저장 및 다운로드
                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)
                
                st.success(f"🎉 `{m3}` 정산 내역서 생성이 완료되었습니다! (청구/미청구 총 {total_records}개 매장 반영)")
                
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
