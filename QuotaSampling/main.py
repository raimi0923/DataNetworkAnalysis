import pandas as pd # 데이터프레임
import numpy as np # 행렬처리
from tkinter import filedialog
from tkinter import messagebox
import tkinter as tk
import tkinter.ttk as ttk
from winreg import *


# 변경할 파일 이름 / 구분 / 크기를 입력받기 위한 코드
root = tk.Tk()

# Gets the requested values of the height and widht.
windowWidth = root.winfo_reqwidth()
windowHeight = root.winfo_reqheight()

# Gets both half the screen width/height and window width/height
positionRight = int(root.winfo_screenwidth() / 2 - windowWidth / 2)
positionDown = int(root.winfo_screenheight() / 2 - windowHeight / 2)

# Positions the window in the center of the page.
root.geometry("+{}+{}".format(positionRight, positionDown))

# 파일 선택하기
root.filename = filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
filename = root.filename

# 구분과 크기 입력 창 만들기
ttk.Label(root, text="구분과 크기를 정하세요").grid(column=0,row=0)

# 구분 리스트
levels = ['광역시도' ,'시군구' ,'읍면동']
level_lst_Chosen = ttk.Combobox(root, width=12, values=levels)
level_lst_Chosen.grid(column=0, row=1)

# 크기 정하기
var = tk.IntVar().set(1000)
num = ttk.Entry(root, width=10, text=var)
num.grid(column=1, row=1)

# 버튼 만들기
def clickMe():
    # 전역 변수 설정하기
    global level
    global num

    level = level_lst_Chosen.get()
    if not(isinstance(int(num.get()), int)):
        messagebox.showerror("메세지박스","크기는 숫자만 입력하세요.")
        exit()
    num = int(num.get())
    root.destroy()

action=ttk.Button(root, text="시작", command=clickMe)
action.grid(column=2, row=1)

# 윈도우 창 실행하기
root.mainloop()

# levels에서 선택된 level은 제외
levels.remove(level)

# Load Data
df = pd.read_excel("{}".format(filename), sheet_name=1)
df = df[df.구분==level]
df = df.drop(levels, axis=1)

# Define Features
male_cols = ['남 19-29세', '남 30대', '남 40대', '남 50대', '남 60세 이상']
female_cols = ['여 19-29세', '여 30대', '여 40대', '여 50대', '여 60세이상']
total_cols = male_cols + female_cols

# Total Population
try:
    total_pop = df[total_cols].sum().sum()
except:
    messagebox.showerror("메세지 박스","해당 파일의 기준 변수명이 다릅니다.")
    exit()

# 2단계 반올림 전
before_df = df.copy()
before_df[total_cols] = (df[total_cols] / total_pop) * num # 각 셀값을 전체 인구로 나누고 정해진 값으로 곱ㅎ
before_df['남 합계'] = before_df[male_cols].sum(axis=1)
before_df['여 합계'] = before_df[female_cols].sum(axis=1)
before_df['총계'] = before_df[['남 합계' ,'여 합계']].sum(axis=1)
# 2단계 남여 각각 합계의 반올림
before_sex_sum = before_df[['남 합계' ,'여 합계']].sum().round()

# 3단계 반올림 후
after_df = df.copy()
after_df[total_cols] = (df[total_cols] / total_pop) * num # 각 셀값을 전체 인구로 나누고 정해진 값으로 곱ㅎ
after_df[total_cols] = after_df[total_cols].round() # 각 셀을 반올림
after_df['남 합계'] = after_df[male_cols].sum(axis=1)
after_df['여 합계'] = after_df[female_cols].sum(axis=1)
after_df['총계'] = after_df[['남 합계' ,'여 합계']].sum(axis=1)
# 3단계 남여 각각 합계의 반올림
after_sex_sum = after_df[['남 합계' ,'여 합계']].sum()

# 2,3단계 남여 합계의 차이
'''
차이는 세 가지 경우로 나뉜다: 남여 각각 차이가 1. 0이거나 / 2. 0보다 크거나 / 3. 0보다 작거나 
1. 0인 경우는 추가적인 일 없이 표 완성
2. 만약 차이가 0보다 큰 경우 : xx.5 보다 작고 xx.5에 가장 가까운 값인 반올림 값에 + 1 
- Why? 반올림하여 내림이 된 값 중 가장 올림에 가까운 값에 1을 더하는 것이 이상적이기 때문
    ex) 2.49999 -> round(2.49999) + 1
3. 만약 차이가 0보다 작은 경우 : xx.5 이상이고 xx.5에 가장 가까운 값인 반올림 값에 - 1
- Why? 반올림하여 올림이 된 값 중 가장 내림에 가까운 값에 1을 빼는 것이 이상적이기 때문
    ex) 2.50001 -> round(2.50001) - 1
'''
sex_diff = before_sex_sum - after_sex_sum

# 성별 합계 차이를 매꾸는 단계
sex_cols_lst = [male_cols, female_cols]
sex_idx = ['남 합계' ,'여 합계']
for i in range(len(sex_idx)):
    if sex_diff.loc[sex_idx[i]] > 0:
        # 차이가 0보다 큰 경우
        '''
        1. 2단계 반올림 전 값을 모두 내림 한 후 0.5를 더 한다.
        2. 1번에서 한 값과 2단계 반올림 전 값을 뺀다.
        3. 음수로 나오는 값은 모두 1로 변환. 1이 가장 큰 값이기 때문.
        ex) 13.45 -> 13 으로 내림 후 (13 + 0.5) - 13.45 = 0.05
        '''
        temp = (before_df[sex_cols_lst[i]].astype(int) + 0.5) - before_df[sex_cols_lst[i]] # 1,2번
        temp = temp[temp >0].fillna(1) # 3번
        v = 1
    elif sex_diff.loc[sex_idx[i]] < 0:
        # 차이가 0보다 작은 경우
        '''
        1. 2단계 반올림 전 값을 모두 내림 한 후 0.5를 더 한다.
        2. 1번에서 한 값과 2단계 반올림 전 값을 뺀다.
        3. 음수로 나오는 값은 모두 1로 변환. 1이 가장 큰 값이기 때문.
        ex) 13.54 -> 13 으로 내림 후 13.54 - (13 + 0.5) = 0.04
        '''
        temp = before_df[sex_cols_lst[i]] - (before_df[sex_cols_lst[i]].astype(int) + 0.5) # 1,2번
        temp = temp[temp >0].fillna(1) # 3번에 해당
        v = -1
    else:
        # 차이가 0인 경우는 이후 과정 생략하고 그냥 통과
        continue

    # 실제합계와의 차이: 절대값을 통해서 음수를 변환하고 정수로 타입을 변환
    cnt = int(abs(sex_diff.loc[sex_idx[i]]))
    row_col = np.unravel_index(np.argsort(temp.values.ravel())[:cnt], temp.shape)
    rows = row_col[0]
    cols = row_col[1]
    # 각 (행,열) 좌표값에 v를 더함
    for r in range(len(rows)):
        temp = after_df[sex_cols_lst[i]].copy()
        temp.iloc[rows[r] ,cols[r]] = temp.iloc[rows[r] ,cols[r]] + v
        after_df[sex_cols_lst[i]] = temp
    print()


# 부족한 부분이 채워졌으면 합계 계산
after_df['남 합계'] = after_df[male_cols].sum(axis=1)
after_df['여 합계'] = after_df[female_cols].sum(axis=1)
after_df['총계'] = after_df[['남 합계' ,'여 합계']].sum(axis=1)
final_sex_sum = after_df[['남 합계' ,'여 합계']].sum()

# 다운로드 폴더 경로 찾기
with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
    Downloads = QueryValueEx(key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]

# 완료 메세지

if final_sex_sum.sum() != num:
    messagebox.showerror("메세지 상자","합계가 0이 아닙니다. 문제를 확인해주세요.")
else:
    messagebox.showinfo("메세지 상자", "다운로드 폴더에 저장되었습니다.")
    save_name = filename.split('/')[-1]
    after_df.to_excel('{}/[{}_{}]{}'.format(Downloads, level, num, save_name), index=False, encoding='cp949')

