import streamlit as st
import pandas as pd
import re
import io
from datetime import datetime

# 페이지 설정
st.set_page_config(page_title="식대 대장 자동 추출기", layout="centered")

# 이름 추출 함수 (기존 로직 동일)
def extract_name(text):
    if pd.isna(text) or text == "": return None
    first_line = str(text).split('\n')[0]
    match = re.search(r'\d+\.(.*)', first_line)
    return match.group(1).strip() if match else None

# 헤더 부분
st.title("🍱 식대 대장 자동 추출기 v1.0")
st.info("엑셀 파일들을 업로드하면 성명(B열)과 식대금액(D열)을 자동으로 정리합니다.")

# 1. 저장 방식 선택
save_mode = st.radio("저장 방식을 선택하세요", ["하나로 통합 저장", "각각 따로 저장"], index=0)

# 2. 파일 업로드 (여러 개 가능)
uploaded_files = st.file_uploader("엑셀 파일을 드래그해서 놓거나 클릭하세요", 
                                  type=["xlsx", "xls"], 
                                  accept_multiple_files=True)

if uploaded_files:
    all_data = []
    processed_files = {} # 각각 저장을 위한 딕셔너리
    now = datetime.now().strftime("%Y%m%d_%H%M")

    for uploaded_file in uploaded_files:
        try:
            df = pd.read_excel(uploaded_file, header=None, engine='openpyxl')
            processed_data = []
            
            for _, row in df.iterrows():
                if len(row) < 4: continue
                name = extract_name(row[1]) # B열
                if name:
                    amount = row[3] if pd.notna(row[3]) else 0 # D열
                    processed_data.append({'성명': name, '식대금액': amount})
            
            if processed_data:
                temp_df = pd.DataFrame(processed_data)
                if save_mode == "하나로 통합 저장":
                    temp_df['출처파일명'] = uploaded_file.name
                    all_data.append(temp_df)
                else:
                    processed_files[uploaded_file.name] = temp_df
        except Exception as e:
            st.error(f"파일 처리 오류 ({uploaded_file.name}): {e}")

    # 3. 결과 다운로드
    st.divider()
    
    if save_mode == "하나로 통합 저장" and all_data:
        merged_df = pd.concat(all_data, ignore_index=True)
        # 메모리 상에서 엑셀 파일 생성
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False)
        
        st.success(f"총 {len(uploaded_files)}개의 파일이 통합되었습니다.")
        st.download_button(
            label="📥 통합 결과 파일 다운로드",
            data=output.getvalue(),
            file_name=f"{now}_식대대장_통합결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    elif save_mode == "각각 따로 저장" and processed_files:
        st.success(f"각 파일별로 정리가 완료되었습니다.")
        for name, df in processed_files.items():
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            clean_name = name.split('.')[0]
            st.download_button(
                label=f"📥 {name} 다운로드",
                data=output.getvalue(),
                file_name=f"{now}_{clean_name}_정리완료.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# 제작자 정보
st.divider()
st.caption(f"제작자: [bamnamoo@gmail.com] | Copyright 2026. All rights reserved.")