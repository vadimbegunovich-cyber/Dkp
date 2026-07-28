import streamlit as st
import docx
from docx import Document
import io

st.set_page_config(page_title="Автоматизация ДКП и Актов", page_icon="🚗", layout="wide")

st.title("🚗 Автоматизация заполнения ДКП и Актов")
st.caption("Автономная версия")

if "form_data" not in st.session_state:
    st.session_state.form_data = {
        "seller_fio": "",
        "seller_passport": "",
        "seller_address": "",
        "buyer_fio": "ООО «Авто-К»",
        "buyer_details": "ИНН 7700000000, ОГРН 1117700000000, г. Москва",
        "car_mark": "",
        "car_vin": "",
        "car_year": "",
        "car_pts": "",
        "car_sts": "",
        "car_number": "",
        "price": "",
        "price_str": ""
    }

st.subheader("📋 Данные для ДКП и Акта")
tab1, tab2, tab3 = st.tabs(["👤 Участники", "🚘 Автомобиль", "💰 Финансы"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("##### Продавец (Физ. лицо)")
        st.session_state.form_data["seller_fio"] = st.text_input("ФИО Продавца", st.session_state.form_data["seller_fio"])
        st.session_state.form_data["seller_passport"] = st.text_area("Паспортные данные", st.session_state.form_data["seller_passport"], height=100)
        st.session_state.form_data["seller_address"] = st.text_input("Адрес регистрации", st.session_state.form_data["seller_address"])
    with col2:
        st.markdown("##### Покупатель")
        st.session_state.form_data["buyer_fio"] = st.text_input("Покупатель / Организация", st.session_state.form_data["buyer_fio"])
        st.session_state.form_data["buyer_details"] = st.text_area("Реквизиты покупателя", st.session_state.form_data["buyer_details"], height=100)

with tab2:
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.form_data["car_mark"] = st.text_input("Марка / Модель ТС", st.session_state.form_data["car_mark"])
        st.session_state.form_data["car_vin"] = st.text_input("VIN номер", st.session_state.form_data["car_vin"])
        st.session_state.form_data["car_year"] = st.text_input("Год выпуска", st.session_state.form_data["car_year"])
    with col2:
        st.session_state.form_data["car_pts"] = st.text_input("Паспорт ТС (ПТС/ЭПТС)", st.session_state.form_data["car_pts"])
        st.session_state.form_data["car_sts"] = st.text_input("Свидетельство ТС (СТС)", st.session_state.form_data["car_sts"])
        st.session_state.form_data["car_number"] = st.text_input("Государственный номер", st.session_state.form_data["car_number"])

with tab3:
    col1, col2 = st.columns(2)
    with col1:
        st.session_state.form_data["price"] = st.text_input("Стоимость (цифрами, руб)", st.session_state.form_data["price"])
    with col2:
        st.session_state.form_data["price_str"] = st.text_input("Стоимость (прописью)", st.session_state.form_data["price_str"])

st.divider()

if st.button("📄 Сформировать полный комплект (ДКП + Акт)", type="primary", use_container_width=True):
    doc = Document()
    
    h = doc.add_heading('ДОГОВОР КУПЛИ-ПРОДАЖИ ТРАНСПОРТНОГО СРЕДСТВА', level=1)
    h.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    d = st.session_state.form_data
    
    doc.add_paragraph(f"Мы, нижеподписавшиеся:\n"
                      f"Продавец: {d['seller_fio']}, паспорт: {d['seller_passport']}, проживающий по адресу: {d['seller_address']},\n"
                      f"и Покупатель: {d['buyer_fio']}, реквизиты: {d['buyer_details']},\n"
                      f"заключили настоящий Договор о нижеследующем:")
    
    doc.add_heading('1. Предмет договора', level=2)
    doc.add_paragraph(f"1.1. Продавец обязуется передать в собственность Покупателя, а Покупатель принять и оплатить транспортное средство:\n"
                      f"• Марка, модель: {d['car_mark']}\n"
                      f"• Идентификационный номер (VIN): {d['car_vin']}\n"
                      f"• Год выпуска: {d['car_year']}\n"
                      f"• ПТС / ЭПТС: {d['car_pts']}\n"
                      f"• СТС: {d['car_sts']}\n"
                      f"• Государственный регистрационный знак: {d['car_number']}")
    
    doc.add_heading('2. Стоимость и порядок расчетов', level=2)
    doc.add_paragraph(f"2.1. Стоимость транспортного средства составляет {d['price']} ({d['price_str']}) рублей.")
    
    doc.add_page_break()
    
    h2 = doc.add_heading('АКТ ПРИЕМА-ПЕРЕДАЧИ ТРАНСПОРТНОГО СРЕДСТВА', level=1)
    h2.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph(f"Настоящий Акт составлен о том, что Продавец ({d['seller_fio']}) передал, "
                      f"а Покупатель ({d['buyer_fio']}) принял транспортное средство {d['car_mark']}, "
                      f"VIN: {d['car_vin']}. Претензий по техническому состоянию и расчетам стороны не имеют.")
    
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    
    st.download_button(
        label="💾 Скачать готовый файл (.docx)",
        data=target,
        file_name=f"ДКП_{d['car_mark']}_{d['car_vin']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
