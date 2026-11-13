# 상담원장 정리 스크립트
# 근거: my-wiki/decisions/d-001-상담원장-정리-규칙.md  (결정 1~5)
#   결정 1  종결 -> 완료, 대기 -> 보류
#   결정 2  처리시간 음수는 지우지 않고 _표시 칸에 이유를 적는다
#   결정 3  만족도 6~10점은 지우지 않고 _표시 칸에 이유를 적고, 평균 만족도에서는 뺀다
#   결정 4  칸 이름은 원본 그대로 (_표시는 맨 끝에 더하는 칸)
#   결정 5  고객번호 앞뒤 공백을 없앤다
# 원본은 읽기만 한다. 쓰는 곳: 정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
# 실행: python3 정리/정리_상담원장.py   (이 폴더 어디서 불러도 된다)
import re
from pathlib import Path
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT.parent / 'messy' / '상담원장_2025.xlsx'
OUT = ROOT / '정리본'
OUT.mkdir(exist_ok=True)

MARK = '_표시'
# 이유 문구는 임시 문구다(결정 기록에서 «미확인»). 숫자에는 영향이 없다.
MARK_NEG = '처리시간 음수: 틀린 값이지만 진짜 값을 몰라 그대로 둠'
MARK_SAT = '만족도 6~10점: NPS를 잘못 넣은 것으로 보이나 확인하지 못함. 평균 만족도에서 뺌'

raw = pd.read_excel(SRC, sheet_name='원장', dtype=object)
cols = list(raw.columns)
raw.insert(0, '원본행', raw.index + 2)
N0 = len(raw)

# ---- 1. 지울 줄: 테스트 / 이전 처리상태 줄 / 중복 추출 ------------------------------
reason = pd.Series('', index=raw.index)
test = raw['상담원ID'] == 'AG999'
reason[test] = '테스트 접수 (상담원ID AG999)'
r = raw[~test]
latest = r.groupby('접수번호')['최종수정일'].transform('max')
older = r['최종수정일'] < latest
reason[older[older].index] = '이전 처리상태 줄 (최종수정일이 늦은 줄을 남김)'
top = r[~older]
dup = top.duplicated('접수번호', keep='last')      # 최종수정일이 같으면 파일에서 뒤에 있는 줄을 남긴다
reason[dup[dup].index] = '중복 추출 (최종수정일이 같은 줄 하나만 남김)'

removed = raw[reason != ''].copy()
removed['지운이유'] = reason[reason != '']
removed.to_csv(OUT / '지운줄.csv', index=False, encoding='utf-8-sig')

df = raw[reason == ''].copy()
orig = df.copy()          # 바뀐 줄 수를 세기 위한 비교용
assert len(df) == N0 - len(removed)
assert df['접수번호'].is_unique

changes = []             # (칸, 무엇을 어떻게, 바뀐 줄 수)


def same(a, b):
    if pd.isna(a) and pd.isna(b):
        return True
    if pd.isna(a) or pd.isna(b):
        return False
    try:
        return float(a) == float(b)
    except (TypeError, ValueError):
        return str(a) == str(b)


def log(col, what, before, after):
    n = int(sum(not same(a, b) for a, b in zip(before, after)))
    changes.append((col, what, n))


# ---- 2. 표기 통일 --------------------------------------------------------------------
def step(col, what, fn):
    before = df[col].copy()
    df[col] = df[col].map(fn)
    log(col, what, before, df[col])


step('접수일', '날짜 표기를 YYYY-MM-DD 하나로', lambda v: re.sub(r'[/.]', '-', v))
step('고객번호', '앞뒤 공백 없앰 (결정 5)', lambda v: v.strip())
CH = {'전화': '전화', '유선': '전화', '전화상담': '전화', 'TEL': '전화',
      '챗봇': '챗봇', 'chatbot': '챗봇', 'Chatbot': '챗봇',
      '앱': '앱', '모바일앱': '앱', 'app': '앱', 'APP': '앱',
      '이메일': '이메일', '메일': '이메일', 'e-mail': '이메일', 'E-mail': '이메일'}
step('접수채널', '17가지 표기를 전화·챗봇·앱·이메일 4가지로', lambda v: CH[v.replace(' ', '')])
step('상담원팀', '팀 이름 안의 공백 없앰', lambda v: v.replace(' ', '').strip())
step('지역', '서울시·서울특별시→서울, 경기도→경기', lambda v: {'서울시': '서울', '서울특별시': '서울', '경기도': '경기'}.get(v, v) if isinstance(v, str) else v)
step('월정액', '`33,000원` 꼴을 숫자 33000으로', lambda v: int(re.sub(r'\D', '', v)) if isinstance(v, str) else v)
ST = {'완료': '완료', '처리완료': '완료', '종결': '완료',
      '진행중': '진행중', '처리중': '진행중', '보류': '보류', '대기': '보류'}
step('처리상태', '완료·처리완료·종결→완료 / 진행중·처리중→진행중 / 보류·대기→보류 (결정 1)', lambda v: ST[v.strip()])
n_jong = int((orig['처리상태'].str.strip() == '종결').sum())
n_dae = int((orig['처리상태'].str.strip() == '대기').sum())

YN = {'Y': 'Y', 'y': 'Y', '예': 'Y', 'TRUE': 'Y', 'N': 'N', 'n': 'N', '아니오': 'N', 'FALSE': 'N'}
step('재문의', 'Y·y·예·TRUE→Y / N·n·아니오·FALSE→N', lambda v: YN[v] if isinstance(v, str) else v)

# ---- 3. 상담시간: 7/1 이후 전화는 초 -> 분 -----------------------------------------------
d = pd.to_datetime(df['접수일'])
is_phone_sec = (df['접수채널'] == '전화') & (d >= '2025-07-01')
t = df['상담시간'].astype(float)
n9999 = int((raw['상담시간'].astype(float) == 9999).sum())        # 원본 전체에서 센다
t = t.where(t != 9999)                    # 9999는 빈칸(테스트 줄에만 있어 정리본에는 남지 않는다)
t = pd.Series(np.where(is_phone_sec, t / 60, t), index=df.index).round(1)
before = df['상담시간'].copy()
df['상담시간'] = t
log('상담시간', '7/1 이후 전화 줄: 초 → 분 (60으로 나눔)', before, df['상담시간'])
n_sec = int(is_phone_sec.sum())

# ---- 4. 만족도: 0·-·미응답은 빈칸, `N점`은 숫자, 6~10점은 값 그대로 (결정 3) ----------------------
def sat(v):
    if isinstance(v, str):
        m = re.match(r'^(\d)점$', v)
        return int(m.group(1)) if m else np.nan   # `-` · `미응답`
    if pd.isna(v) or v == 0:
        return np.nan
    return v


before = df['만족도'].copy()
df['만족도'] = df['만족도'].map(sat)
log('만족도', '0·`-`·미응답→빈칸, `5점` 꼴→숫자 (6~10점은 그대로 둠)', before, df['만족도'])

# ---- 5. _표시 칸 (결정 2·3) -------------------------------------------------------------
neg = pd.to_numeric(df['처리시간']) < 0
hi = pd.to_numeric(df['만족도']) > 5
assert not (neg & hi).any()
df[MARK] = ''
df.loc[neg, MARK] = MARK_NEG
df.loc[hi, MARK] = MARK_SAT
changes.append((MARK, '처리시간 음수 — 값은 그대로, 이유만 적음 (결정 2)', int(neg.sum())))
changes.append((MARK, '만족도 6~10점 — 값은 그대로, 이유만 적음 (결정 3)', int(hi.sum())))

# ---- 6. 저장 -----------------------------------------------------------------------------
final_cols = cols + [MARK]
assert final_cols[:len(cols)] == cols          # 결정 4: 원본 칸 이름 그대로
out = df[final_cols]
out.to_csv(OUT / '상담원장_정리.csv', index=False, encoding='utf-8-sig')

# ---- 7. 다섯 숫자 · 고객 수 -----------------------------------------------------------------
yn = out['재문의'].dropna()
sat_ok = pd.to_numeric(out['만족도'])[~hi]
metrics = {
    '상담 건수': len(out),
    '평균 상담시간(분)': round(float(out['상담시간'].mean()), 2),
    '재문의율(%)': round(float((yn == 'Y').mean() * 100), 2),
    '평균 만족도': round(float(sat_ok.mean()), 2),
    '완료율(%)': round(float((out['처리상태'] == '완료').mean() * 100), 2),
    '고객 수': int(out['고객번호'].nunique()),
}
extra = {'재문의 계산 건수': len(yn), '만족도 응답(표시 줄 제외)': int(sat_ok.notna().sum()),
         '만족도 값이 있는 줄(표시 줄 포함)': int(pd.to_numeric(out['만족도']).notna().sum())}

# ---- 8. 변경요약.md ----------------------------------------------------------------------
rc = removed['지운이유'].value_counts()
L = []
L.append('# 변경 요약 — 상담원장_2025.xlsx → 정리본/상담원장_정리.csv\n')
L.append('- 근거: `my-wiki/decisions/d-001-상담원장-정리-규칙.md` (결정 1~5)')
L.append('- 스크립트: `정리/정리_상담원장.py` · 원본은 읽기만 했습니다.\n')
L.append('## 줄 수\n')
L.append('| 단계 | 줄 수 |\n|---|---|')
L.append(f'| 원본 (원장 시트) | {N0:,} |')
for k in ['테스트 접수 (상담원ID AG999)', '이전 처리상태 줄 (최종수정일이 늦은 줄을 남김)', '중복 추출 (최종수정일이 같은 줄 하나만 남김)']:
    L.append(f'| 지운 줄: {k} | −{int(rc.get(k, 0)):,} |')
L.append(f'| 정리본 | {len(out):,} |')
L.append(f'\n{N0:,} − {len(removed):,} = {N0 - len(removed):,}. 지운 줄은 `정리본/지운줄.csv`에 있습니다.\n')
L.append('## 칸마다 바꾼 것\n')
L.append('«바뀐 줄 수»는 정리본에 남은 줄 중 원본 값과 달라진 줄입니다. 칸 이름은 바꾸지 않았고(결정 4), 맨 끝에 `_표시` 칸 하나를 더했습니다.\n')
L.append('| 칸 | 무엇을 어떻게 | 바뀐 줄 수 |\n|---|---|---|')
for c, w, n in changes:
    L.append(f'| {c} | {w} | {n:,} |')
L.append('')
L.append('## 따로 적어 둘 것\n')
L.append(f'- **상담시간**: 2025-07-01 이후 접수된 전화 채널 {n_sec:,}줄은 원본에서 **초 단위**였습니다. 작성기준(전화 시스템 교체)에 따라 60으로 나눠 **분 단위**로 바꿨고 소수 한 자리까지 적었습니다. 나머지 줄은 원래 분 단위입니다.')
L.append(f'- 결정 1: `종결` {n_jong:,}줄이 `완료`로, `대기` {n_dae:,}줄이 `보류`로 합쳐졌습니다.')
L.append(f'- 결정 2·3: `_표시` 칸이 채워진 줄은 {int(neg.sum()) + int(hi.sum())}줄(처리시간 음수 {int(neg.sum())} · 만족도 6~10점 {int(hi.sum())})입니다. 두 칸의 값은 지우거나 고치지 않았습니다. 이 칸의 이유 문구는 임시 문구입니다.')
L.append(f'- 결정 5: 고객번호 앞 공백 {next(n for c, w, n in changes if c == "고객번호"):,}줄을 없앴습니다.')
L.append(f'- 상담시간 9999는 원본에 {n9999}건 있었고 모두 지운 줄(테스트 접수)이라 정리본에는 남지 않았습니다.')
(OUT / '변경요약.md').write_text('\n'.join(L) + '\n', encoding='utf-8')

print('줄:', N0, '→', len(out), '(지운 줄', len(removed), dict(rc), ')')
for k, v in {**metrics, **extra}.items():
    print(f'{k}: {v}')
