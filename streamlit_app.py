import os, io, tempfile
import streamlit as st

# Streamlit Secrets -> environment variables used by the existing app.py
if "OPENAI_API_KEY" in st.secrets:
    os.environ["OPENAI_API_KEY"] = st.secrets["OPENAI_API_KEY"]
if "OPENAI_MODEL" in st.secrets:
    os.environ["OPENAI_MODEL"] = st.secrets["OPENAI_MODEL"]

import app as core

st.set_page_config(page_title="Reading Vocabulary Maker", page_icon="📘", layout="wide")
st.title("📘 Reading Vocabulary Maker")
st.caption("독해 복습형 단어장 · 단어 시험지 · 정답 해설지 생성기")

class UploadAdapter:
    def __init__(self, uploaded):
        self._uploaded = uploaded
        self.filename = uploaded.name
        self.mimetype = uploaded.type or ""
    def read(self):
        return self._uploaded.getvalue()


def docx_bytes(data, kind):
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        path = tmp.name
    try:
        if kind == "vocab":
            core.save_vocab_docx(data, path)
        elif kind == "test":
            core.save_test_docx(data, path)
        else:
            core.save_answer_docx(data, path)
        with open(path, "rb") as f:
            return f.read()
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


def shuffle_questions(data):
    by = {1: [], 2: [], 3: [], 4: []}
    for q in data["questions"]:
        by[q["section"]].append(q)
    import random
    rng = random.SystemRandom()
    qs = []
    for s in (1,2,3,4):
        rng.shuffle(by[s])
        qs += by[s]
    for i, q in enumerate(qs, 1):
        q["number"] = i
    data["questions"] = qs
    return data

with st.container(border=True):
    c1, c2, c3 = st.columns([2.2, 1.2, 1])
    with c1:
        title = st.text_input("시험지 제목", placeholder="예: Reading Inside - Unit 1 Origins")
    with c2:
        page_range = st.text_input("PDF 페이지 범위", placeholder="예: 94-113 또는 94,96-100")
    with c3:
        difficulty_label = st.selectbox("난이도", ["더 쉽게", "기본", "더 어렵게"], index=1)

    files = st.file_uploader(
        "독해 자료 PDF / 이미지",
        type=["pdf", "jpg", "jpeg", "png", "webp"],
        accept_multiple_files=True,
    )

D = {"더 쉽게":"easier", "기본":"normal", "더 어렵게":"harder"}

if st.button("분석 및 생성", type="primary", use_container_width=True):
    if not files:
        st.warning("PDF 또는 이미지 파일을 하나 이상 올려 주세요.")
    else:
        with st.spinner("지문을 분석하고 단어장과 시험지를 만드는 중입니다..."):
            try:
                adapted = [UploadAdapter(f) for f in files]
                st.session_state.result = core.generate(title.strip(), D[difficulty_label], adapted, page_range.strip())
                st.success("생성이 완료되었습니다.")
            except Exception as e:
                st.exception(e)

if "result" in st.session_state:
    data = st.session_state.result
    st.info(f"추정 난이도: {data.get('detected_level','—')}" + (f" | {data.get('source_note')}" if data.get('source_note') else ""))

    if st.button("시험 순서 다시 섞기"):
        st.session_state.result = shuffle_questions(data)
        st.rerun()

    t1, t2, t3 = st.tabs(["단어장", "시험지", "정답·해설"])

    with t1:
        rows = []
        for i, v in enumerate(data["vocabulary"], 1):
            rows.append({
                "No.": i, "Word": v["word"], "발음": v["pronunciation"], "품사": v["pos"],
                "뜻": v["meaning_ko"], "Synonym": v["synonym"], "Antonym": v["antonym"],
                "New Example": v["example"],
            })
        st.dataframe(rows, use_container_width=True, hide_index=True)

    with t2:
        sec = None
        for q in data["questions"]:
            if q["section"] != sec:
                sec = q["section"]
                st.subheader(str(sec))
            st.markdown(f"**{q['number']}.** {q['prompt']}")
            if q["choices"]:
                st.write("　".join(q["choices"]))
            st.divider()

    with t3:
        rows = [{"No.":q["number"], "정답":q["answer"], "뜻":q["meaning_ko"], "문제 해설":q["explanation_ko"]} for q in data["questions"]]
        st.dataframe(rows, use_container_width=True, hide_index=True)

    st.subheader("파일 다운로드")
    b1, b2, b3 = st.columns(3)
    with b1:
        st.download_button("단어장 Word", docx_bytes(data,"vocab"), "Vocabulary_List.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with b2:
        st.download_button("시험지 Word", docx_bytes(data,"test"), "Vocabulary_Test.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with b3:
        st.download_button("정답·해설 Word", docx_bytes(data,"answer"), "Answer_Explanation.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
