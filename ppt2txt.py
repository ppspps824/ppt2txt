import os
import tempfile
import time
from pathlib import Path

import openai
import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE


def think_answer(text, model):

    prompt = f"""
あなたはベテランのプレゼンテーターです。

以下文章について3点コメントをお願いします。

1.いい点

具体的かつ最大限褒めちぎる。

2.改善できる点

プレゼンテーションの文章として最も重要かつ簡単に取り組める内容で簡潔に一つ。

3.質問

聴講者の理解が深まる内容で簡潔に一つ。

{text}
    
    """

    messages = [{"role": "user", "content": prompt}]
    try_count = 3
    error_mes = ""
    for try_time in range(try_count):
        try:
            resp = openai.ChatCompletion.create(
                model=model,
                messages=messages,
                stream=True,
                timeout=120,
                request_timeout=120,
            )
            return resp

        except openai.error.APIError as e:
            print(e)
            print(f"retry:{try_time+1}/{try_count}")
            error_mes = e
            time.sleep(1)
        except openai.error.InvalidRequestError as e:
            print(e)
            print(f"retry:{try_time+1}/{try_count}")
            error_mes = e
            pass
        except (
            openai.error.RateLimitError,
            openai.error.openai.error.APIConnectionError,
        ) as e:
            print(e)
            print(f"retry:{try_time+1}/{try_count}")
            error_mes = e
            time.sleep(10)

    st.error(error_mes)
    st.stop()


# スライドから文字を抽出
def check_recursively_for_text(this_set_of_shapes, txt_list):
    for shape in order_shapes(this_set_of_shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            check_recursively_for_text(shape.shapes, txt_list)
        else:
            if hasattr(shape, "text"):
                if shape.text:
                    txt_list.append(f"{shape.text}\n")

            if shape.has_table:  # 表に含まれるテキストを抽出
                for cell in shape.table.iter_cells():
                    txt_list.append(cell.text)

    return txt_list


def order_shapes(shapes):
    return sorted(shapes, key=lambda x: (x.top, x.left))


if __name__ == "__main__":

    os.environ["OPENAI_API_KEY"] = st.secrets["OPEN_AI_KEY"]
    openai.api_key = st.secrets["OPEN_AI_KEY"]

    with st.sidebar:
        file = st.file_uploader("")

    if file:
        with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
            fp = Path(tmp_file.name)
            fp.write_bytes(file.getvalue())

        # ファイルの読み込み
        prs = Presentation(fp)

        # スライドごとに繰り返し
        all_txt_list = []
        for num, slide in enumerate(prs.slides):
            # テキスト抽出
            all_txt_list.append(f"## slide{num+1}\n")
            all_txt_list = check_recursively_for_text(slide.shapes, all_txt_list)

            # 発表者ノートからテキスト抽出
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame.text:
                all_txt_list.append(slide.notes_slide.notes_text_frame.text)

        all_text = "\n".join(all_txt_list)
        with st.expander(f"抽出結果:{file.name}"):
            st.code(all_text)

        col1, col2 = st.columns(2)

        with col1:
            model = st.selectbox("評価モデルを選択", ["gpt-4", "gpt-3.5-turbo"])
        with col2:
            st.write("")
            st.write("")
            ask_submit = st.button("評価")

        if all([ask_submit, all_text]):
            message_placeholder = st.empty()
            full_response = ""
            for response in think_answer(all_text, model):
                full_response += response["choices"][0]["delta"].get("content", "")
                message_placeholder.write(full_response)
