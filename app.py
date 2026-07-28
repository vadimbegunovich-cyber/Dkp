import streamlit as st
import docx
from docx import Document
import google.generativeai as genai
from PIL import Image
import json
import io

st.set_page_config(page_title="Автоматизация ДКП и Актов", page_icon="🚗", layout="wide")

st.title("🚗 Автоматизация заполнения ДКП и Актов")
st.caption("Финальная версия с распознаванием документов")

st.sidebar.header("⚙️ Настройки и Загрузка")
api_key = st.sidebar.text_input("API-ключ Gemini (для распознавания)", type="password")

if api_key:
    genai.configure(api_key=api_key)

uploaded_files = st.sidebar.file_uploader(
    "Загрузите фото документов (Паспорт, СТС, ПТС)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

if "form_data" not in st.session_state:
    st.session_state.form_data = {
        "seller_fio": "",
        "seller_passport": "",
        "seller_address": "",
        "buyer_fio": "ООО «Авто-К»",
        "buyer_details": "",
        "car_mark": "",
        "car_vin": "",
        "car_year": "",
        "car_pts": "",
        "car_sts": "",
        "car_number": "",
        "price": "",
        "price_str": ""
    }

if st.sidebar.button("🤖 Распознать фото", type="primary"):
    if not api_key:
        st.sidebar.error("Введите API-ключ!")
    elif not uploaded_files:
        st.sidebar.warning("Загрузите фото документов.")
    else:
        with st.spinner("Анализ..."):
            try:
                model = genai.GenerativeModel('gemini-1.5-flash')
                images = [Image.open(f) for f in uploaded_files]
                prompt = """Извлеки из фото (паспорт, СТС, ПТС) данные в JSON:
                seller_fio, seller_passport, seller_address, car_mark, car_vin, car_year, car_pts, car_sts, car_number.
                Верни ТОЛЬКО чистый JSON."""
                response = model.generate_content([prompt, *images])
                cleaned_text = response.text.replace("```json", "").replace("```", "").strip()
                extracted_data = json.loads(cleaned_text)
                for key, val in extracted_data.items():
                    if key in st.session_state.form_data and val:
                        st.session_state.form_data[key] = str(val)
                st.sidebar.success("Готово!")
            except Exception as e:
                st.sidebar.error(f"Ошибка: {e}")

st.subheader("📋 Данные для ДКП")
tab1, tab2, tab3 = st.tabs(["👤 Участники", "🚘 Автомобиль", "💰 Финансы"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.form_data["seller_fio"] = st.text_input("ФИО Продавца", st.session_state.form_data["seller_fio"])
        st.session_state.form_data["seller_passport"] = st.text_area("Паспортные данные", st.session_state.form_data["seller_passport"])
        st.session_state.form_data["seller_address"] = st.text_input("Адрес", st.session_state.form_data["seller_address"])
    with col2:
        st.session_state.form_data["buyer_fio"] = st.text_input("Покупатель", st.session_state.form_data["buyer_fio"])
        st.session_state.form_data["buyer_details"] = st.text_area("Реквизиты покупателя", st.session_state.form_data["buyer_details"])

with tab2:
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.form_data["car_mark"] = st.text_input("Марка / Модель", st.session_state.form_data["car_mark"])
        st.session_state.form_data["car_vin"] = st.text_input("VIN", st.session_state.form_data["car_vin"])
        st.session_state.form_data["car_year"] = st.text_input("Год", st.session_state.form_data["car_year"])
    with col2:
        st.session_state.form_data["car_pts"] = st.text_input("ПТС / ЭПТС", st.session_state.form_data["car_pts"])
        st.session_state.form_data["car_sts"] = st.text_input("СТС", st.session_state.form_data["car_sts"])
        st.session_state.form_data["car_number"] = st.text_input("Гос. номер", st.session_state.form_data["car_number"])

with tab3:
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.form_data["price"] = st.text_input("Цена (число)", st.session_state.form_data["price"])
    with col2:
        st.session_state.form_data["price_str"] = st.text_input("Цена (прописью)", st.session_state.form_data["price_str"])

if st.button("📄 Сформировать ДКП (.docx)", type="primary"):
    doc = Document()
    doc.add_heading('ДОГОВОР КУПЛИ-ПРОДАЖИ ТРАНСПОРТНОГО СРЕДСТВА', level=1)
    d = st.session_state.form_data
    doc.add_paragraph(f"Продавец: {d['seller_fio']}, паспорт: {d['seller_passport']}, прож.: {d['seller_address']}.")
    doc.add_paragraph(f"Покупатель: {d['buyer_fio']}, {d['buyer_details']}.")
    doc.add_paragraph(f"Марка: {d['car_mark']}, VIN: {d['car_vin']}, Год: {d['car_year']}, ПТС: {d['car_pts']}, СТС: {d['car_sts']}, Госномер: {d['car_number']}.")
    doc.add_paragraph(f"Стоимость: {d['price']} ({d['price_str']}) руб.")
    
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    
    st.download_button(
        label="💾 Скачать готовый Word-файл",
        data=target,
        file_name=f"ДКП_{d['car_mark']}_{d['car_vin']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
