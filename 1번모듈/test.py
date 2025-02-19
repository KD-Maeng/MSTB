#  데이터 검증을 자동화하기 위해 pandas를 사용하여 입력 파일과 출력 파일을 비교하는 스크립트
#  다음은 이를 구현한 예제입니다:


# 이 스크립트는 다음과 같은 기능을 수행합니다:
# 입력 파일과 출력 파일을 읽어옵니다.
# 출력 파일의 모든 시트를 하나의 데이터프레임으로 합칩니다.
# 누락된 인원이 있는지 확인합니다.
# 등급별로 균등하게 배치되었는지 확인합니다.
# 성별로 균등하게 배치되었는지 확인합니다.
# 이 스크립트를 실행하면 누락된 인원, 등급별 균등 배치, 성별 균등 배치 여부를 자동으로 검증할 수 있습니다.

import pandas as pd

# 파일 경로
input_file_path = r'C:\Users\kdmae\OneDrive\바탕 화면\MSTB\INPUT.xlsx'
output_file_path = r'C:\Users\kdmae\OneDrive\바탕 화면\MSTB\OUTPUT1.xlsx'

# 입력 파일 읽기
input_df = pd.read_excel(input_file_path)

# 출력 파일 읽기
output_dfs = pd.read_excel(output_file_path, sheet_name=None)

# 출력 파일의 모든 시트를 하나의 데이터프레임으로 합치기
output_df = pd.concat(output_dfs.values(), ignore_index=True)

# 누락된 인원 확인
missing_in_output = input_df[~input_df.isin(output_df)].dropna(how='all')
missing_in_input = output_df[~output_df.isin(input_df)].dropna(how='all')

if missing_in_output.empty and missing_in_input.empty:
    print("모든 인원이 정상적으로 배치되었습니다.")
else:
    print("누락된 인원이 있습니다.")
    if not missing_in_output.empty:
        print("출력 파일에 누락된 인원:")
        print(missing_in_output)
    if not missing_in_input.empty:
        print("입력 파일에 누락된 인원:")
        print(missing_in_input)

# 등급별 균등 확인
input_grade_counts = input_df['통합평균등급'].value_counts()
output_grade_counts = output_df['통합평균등급'].value_counts()

if input_grade_counts.equals(output_grade_counts):
    print("등급별로 균등하게 배치되었습니다.")
else:
    print("등급별로 균등하게 배치되지 않았습니다.")
    print("입력 파일 등급 분포:")
    print(input_grade_counts)
    print("출력 파일 등급 분포:")
    print(output_grade_counts)

# 성별 균등 확인
input_gender_counts = input_df['성별'].value_counts()
output_gender_counts = output_df['성별'].value_counts()

if input_gender_counts.equals(output_gender_counts):
    print("성별로 균등하게 배치되었습니다.")
else:
    print("성별로 균등하게 배치되지 않았습니다.")
    print("입력 파일 성별 분포:")
    print(input_gender_counts)
    print("출력 파일 성별 분포:")
    print(output_gender_counts)