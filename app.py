import streamlit as st
import pandas as pd
import os
from io import BytesIO

st.set_page_config(page_title="디버 스멤구분 정산기", layout="centered")

st.title("💰 디버 5월 정산 프로그램 (스멤구분)")
st.write("배차리스트 파일들과 디버 사용내역 파일을 업로드한 후 '정산 시작'을 눌러주세요.")

st.markdown("---")

# 1. 5월 폴더 내의 배차리스트 파일들 업로드 (여러 개 다중 선택 가능)
st.subheader("1. 5월 폴더 안의 배차리스트 엑셀 파일들")
dispatch_files = st.file_uploader(
    "5월 폴더 안에 있던 모든 배차리스트 엑셀 파일들을 한 번에 선택해서 올려주세요.",
    type=["xlsx"],
    accept_multiple_files=True
)

# 2. 디버 사용내역 파일 업로드 (1개)
st.subheader("2. 디버 사용내역 파일")
usage_file = st.file_uploader(
    "디버_사용내역_2026-05-01_2026-05-31.xlsx 파일을 올려주세요.",
    type=["xlsx"]
)

st.markdown("---")

# 두 가지 파일이 모두 최소한으로 업로드 되었을 때 버튼 활성화
if dispatch_files and usage_file is not None:
    st.success(f"배차리스트 {len(dispatch_files)}개 및 사용내역 파일 로드 완료!")

    if st.button("🚀 정산 시작 및 결과 만들기"):
        with st.spinner("기존 정산 로직을 실행 중입니다... 잠시만 기다려주세요."):
            try:
                # [로직 1] 다중 업로드된 배차리스트 파일 합치기
                dispatch_list = []
                for file in dispatch_files:
                    # 임시 파일 이름 확인 (~$ 임시파일 제외)
                    if file.name.startswith("~$"):
                        continue

                    df = pd.read_excel(file)
                    df.columns = df.columns.str.strip()

                    if "디버오더번호" in df.columns:
                        df["디버오더번호"] = pd.to_numeric(df["디버오더번호"], errors="coerce")
                        df = df[df["디버오더번호"].notna()].copy()
                        df["디버오더번호"] = df["디버오더번호"].astype("Int64")
                        dispatch_list.append(df)

                if not dispatch_list:
                    st.error("올바른 '디버오더번호'가 포함된 배차리스트 엑셀 파일이 없습니다.")
                    st.stop()

                dispatch_df = pd.concat(dispatch_list, ignore_index=True)

                # 받는사람 '스콘' 필터링
                arrive_scone_df = dispatch_df[
                    dispatch_df["받는사람"].astype(str).str.contains("스콘", na=False)
                ].copy()

                scone_order_numbers = set(
                    arrive_scone_df["디버오더번호"].dropna().astype("Int64")
                )

                # [로직 2] 사용내역 파일 처리
                usage_df = pd.read_excel(usage_file)
                usage_df.columns = usage_df.columns.str.strip()

                usage_df["주문번호"] = pd.to_numeric(usage_df["주문번호"], errors="coerce")
                usage_df["경유지개수"] = pd.to_numeric(usage_df["경유지개수"], errors="coerce").fillna(0)
                usage_df["결제금액"] = pd.to_numeric(usage_df["결제금액"], errors="coerce").fillna(0)

                start_scone_df = usage_df[
                    usage_df["발송인"].astype(str).str.contains("스콘", na=False)
                ].copy()

                # [로직 3] 스멤구분 및 금액 쪼개기 계산
                result_rows = []
                for _, row in usage_df.iterrows():
                    row_data = row.copy()

                    order_no = row_data["주문번호"]
                    sender = str(row_data.get("발송인", ""))
                    stop_count = row_data["경유지개수"]
                    payment = row_data["결제금액"]

                    is_membership = (
                            order_no in scone_order_numbers or "스콘" in sender
                    )

                    if is_membership:
                        if stop_count >= 1:
                            membership_amount = payment / (stop_count + 1)
                            studio_amount = payment - membership_amount

                            member_row = row_data.copy()
                            member_row["스/멤구분"] = "멤버십"
                            member_row["총 금액"] = membership_amount
                            result_rows.append(member_row)

                            studio_row = row_data.copy()
                            studio_row["스/멤구분"] = "스튜디오"
                            studio_row["총 금액"] = studio_amount
                            result_rows.append(studio_row)
                        else:
                            row_data["스/멤구분"] = "멤버십"
                            row_data["총 금액"] = payment
                            result_rows.append(row_data)
                    else:
                        row_data["스/멤구분"] = "스튜디오"
                        row_data["총 금액"] = payment
                        result_rows.append(row_data)

                result_df = pd.DataFrame(result_rows)
                result_df["총 금액"] = result_df["총 금액"].astype(float)

                # 요약 데이터 생성
                membership_total = result_df.loc[result_df["스/멤구분"] == "멤버십", "총 금액"].sum()
                studio_total = result_df.loc[result_df["스/멤구분"] == "스튜디오", "총 금액"].sum()

                summary_df = pd.DataFrame({
                    "항목": ["멤버십", "스튜디오", "전체합계"],
                    "금액": [membership_total, studio_total, membership_total + studio_total]
                })

                # [로직 4] 메모리 버퍼에 멀티 시트 엑셀 파일 작성
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    usage_df.to_excel(writer, sheet_name="원본", index=False)
                    start_scone_df.to_excel(writer, sheet_name="출발지_스콘", index=False)
                    arrive_scone_df.to_excel(writer, sheet_name="도착 및 경유_스콘", index=False)
                    result_df.to_excel(writer, sheet_name="스멤구분", index=False)
                    summary_df.to_excel(writer, sheet_name="요약", index=False)
                processed_data = output.getvalue()

                st.balloons()
                st.success("✨ 정산 가공 완료!")

                # 다운로드 버튼 출력
                st.download_button(
                    label="📥 최종_5월_정산_디버_스멤구분.xlsx 다운로드",
                    data=processed_data,
                    file_name="최종_5월_정산_디버_스멤구분.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"정산 도중 에러가 발생했습니다: {e}")
else:
    st.info("💡 두 영역에 파일을 모두 추가하시면 '정산 시작' 버튼이 나타납니다.")
