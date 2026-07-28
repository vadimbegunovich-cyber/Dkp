import streamlit as st
import docx
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import requests
import re
from PIL import Image
import datetime

st.set_page_config(page_title="Автоматизация ДКП и Актов", page_icon="🚗", layout="wide")

st.title("🚗 Генератор ДКП и Актов (по шаблону)")
st.caption("Автоматическое сканирование фото (открытый API OCR) и формирование полного комплекта документов")

# --- БОКОВАЯ ПАНЕЛЬ: ЗАГРУЗКА И OCR ---
st.sidebar.header("📁 Загрузка документов")

uploaded_files = st.sidebar.file_uploader(
    "Фото (Паспорт, СТС, ПТС)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

def prepare_image(file_bytes):
    """Сжимаем фото до ~100-300 КБ, чтобы бесплатный OCR не выдавал таймаут"""
    img = Image.open(io.BytesIO(file_bytes))
    img.thumbnail((1500, 1500))
    buffer = io.BytesIO()
    img.convert("RGB").save(buffer, format="JPEG", quality=75)
    return buffer.getvalue()

# --- ИНИЦИАЛИЗАЦИЯ ДАННЫХ ---
today_str = datetime.datetime.now().strftime("%d.%m.%Y")

if "form_data" not in st.session_state:
    st.session_state.form_data = {
        # Сделка
        "contract_num": "АК00000000",
        "contract_date": f"{today_str} г.",
        "city": "г. Калининград",
        
        # Покупатель
        "buyer_company": "ООО «Авто-К»",
        "buyer_rep": "Менеджера по закупкам автомобилей с пробегом Марков Андрей Дмитриевич",
        "buyer_poa": "Доверенности б/н от 10.06.2026",
        "buyer_address": "236003, Калининградская обл., г. Калининград, Московский пр-кт, д. 250, помещение 1",
        "buyer_phone": "8 (4012) 777-213",
        "buyer_inn": "3906203848",
        "buyer_kpp": "390601001",
        "buyer_ogrn": "1083925041310",
        "buyer_bank": "КБ \"ЭНЕРГОТРАНСБАНК\" (АО)",
        "buyer_rs": "40702810800000011431",
        "buyer_ks": "30101810800000000701",
        "buyer_bik": "042748701",
        
        # Продавец
        "seller_fio": "",
        "seller_passport": "",
        "seller_address": "",
        "seller_phone": "",
        
        # Авто
        "car_mark": "",
        "car_vin": "",
        "car_engine": "Отсутствует",
        "car_body": "Отсутствует",
        "car_frame": "Отсутствует",
        "car_color": "",
        "car_mileage": "",
        "car_pts": "",
        "car_sts": "",
        "car_number": "",
        "car_year": "",
        "car_defects": "диагностика прилагается",
        
        # Финансы
        "price_num": "1 801 000",
        "price_text": "Один миллион восемьсот одна тысяча рублей 00 копеек",
        "payment_details": "наличным/безналичным путем"
    }

# --- КНОПКА СКАНИРОВАНИЯ ---
if st.sidebar.button("🤖 Распознать фото", type="primary"):
    if not uploaded_files:
        st.sidebar.warning("Загрузите фото для распознавания.")
    else:
        with st.spinner("Сжимаем и отправляем на открытый API OCR..."):
            full_text = ""
            for file in uploaded_files:
                try:
                    compressed_bytes = prepare_image(file.getvalue())
                    # Используем открытый ключ OCR.space (безлимит на сжатые фото)
                    response = requests.post(
                        "https://api.ocr.space/parse/image",
                        files={"filename": (file.name, compressed_bytes, "image/jpeg")},
                        data={
                            "apikey": "K87889882388957",
                            "language": "rus",
                            "isOverlayRequired": False,
                            "scale": True,
                            "OCREngine": "2"
                        },
                        timeout=40
                    )
                    res_json = response.json()
                    if res_json.get("IsErroredOnProcessing"):
                        st.sidebar.error(f"Ошибка API: {res_json.get('ErrorMessage')}")
                    else:
                        for parsed in res_json.get("ParsedResults", []):
                            full_text += "\n" + parsed.get("ParsedText", "")
                except Exception as e:
                    st.sidebar.error(f"Сбой сети: {e}")

            if full_text.strip():
                st.sidebar.success("Текст распознан! Извлекаем данные...")
                full_text_upper = full_text.upper()
                
                # Поиск VIN (17 символов)
                vin_match = re.search(r'\b[A-HJ-NPR-Z0-9]{17}\b', full_text_upper)
                if vin_match:
                    st.session_state.form_data["car_vin"] = vin_match.group(0)
                    st.session_state.form_data["car_body"] = vin_match.group(0) # Часто номер кузова = VIN

                # Гос. номер (А 123 АА 39)
                number_match = re.search(r'[А-ЯA-Z]\s*\d{3}\s*[А-ЯA-Z]{2}\s*\d{2,3}', full_text_upper)
                if number_match:
                    st.session_state.form_data["car_number"] = number_match.group(0).replace(" ", "")

                # Год выпуска
                year_match = re.search(r'\b(19|20)\d{2}\b', full_text)
                if year_match:
                    st.session_state.form_data["car_year"] = year_match.group(0)

                # Серия и номер паспорта (4 цифры, пробел, 6 цифр)
                passport_match = re.search(r'\b\d{2}\s*\d{2}\s*\d{6}\b', full_text)
                if passport_match:
                    st.session_state.form_data["seller_passport"] = passport_match.group(0)

                # ПТС / СТС серии
                docs_match = re.findall(r'\b\d{2}\s*[А-ЯA-Z0-9]{2}\s*\d{6}\b', full_text_upper)
                if len(docs_match) > 0:
                    st.session_state.form_data["car_sts"] = docs_match[0]
                if len(docs_match) > 1:
                    st.session_state.form_data["car_pts"] = docs_match[1]

                with st.sidebar.expander("🔍 Посмотреть сырой текст"):
                    st.text(full_text)
                st.rerun()
            else:
                st.sidebar.error("Текст не найден. Попробуйте фото лучшего качества.")

# --- ИНТЕРФЕЙС ВВОДА ДАННЫХ ---
d = st.session_state.form_data

st.subheader("📋 Редактор данных")
tab_deal, tab_seller, tab_buyer, tab_car = st.tabs(["📄 Договор и Сумма", "👤 Продавец", "🏢 Покупатель", "🚘 Автомобиль"])

with tab_deal:
    c1, c2, c3 = st.columns(3)
    d["contract_num"] = c1.text_input("Номер Договора", d["contract_num"])
    d["contract_date"] = c2.text_input("Дата", d["contract_date"])
    d["city"] = c3.text_input("Город", d["city"])
    
    c4, c5 = st.columns([1, 2])
    d["price_num"] = c4.text_input("Сумма (цифрами)", d["price_num"])
    d["price_text"] = c5.text_input("Сумма (прописью)", d["price_text"])
    d["payment_details"] = st.text_area("Особые условия оплаты (п. 2.2)", d["payment_details"], height=80)

with tab_seller:
    d["seller_fio"] = st.text_input("ФИО Продавца", d["seller_fio"])
    d["seller_passport"] = st.text_input("Паспорт (серия, номер, кем/когда выдан)", d["seller_passport"])
    d["seller_address"] = st.text_input("Адрес регистрации", d["seller_address"])
    d["seller_phone"] = st.text_input("Телефон Продавца", d["seller_phone"])

with tab_buyer:
    c1, c2 = st.columns(2)
    d["buyer_company"] = c1.text_input("Название организации (Покупатель)", d["buyer_company"])
    d["buyer_rep"] = c1.text_area("В лице (Должность, ФИО)", d["buyer_rep"])
    d["buyer_poa"] = c1.text_input("Действующего на основании", d["buyer_poa"])
    
    d["buyer_address"] = c2.text_area("Юр. Адрес", d["buyer_address"])
    d["buyer_inn"] = c2.text_input("ИНН", d["buyer_inn"])
    d["buyer_kpp"] = c2.text_input("КПП", d["buyer_kpp"])
    d["buyer_ogrn"] = c2.text_input("ОГРН", d["buyer_ogrn"])
    
    c3, c4 = st.columns(2)
    d["buyer_bank"] = c3.text_input("Банк", d["buyer_bank"])
    d["buyer_rs"] = c3.text_input("Р/С", d["buyer_rs"])
    d["buyer_ks"] = c4.text_input("К/С", d["buyer_ks"])
    d["buyer_bik"] = c4.text_input("БИК", d["buyer_bik"])

with tab_car:
    c1, c2, c3 = st.columns(3)
    d["car_mark"] = c1.text_input("Марка (Модель)", d["car_mark"])
    d["car_vin"] = c2.text_input("VIN", d["car_vin"])
    d["car_year"] = c3.text_input("Год выпуска", d["car_year"])
    
    c4, c5, c6 = st.columns(3)
    d["car_color"] = c4.text_input("Цвет", d["car_color"])
    d["car_number"] = c5.text_input("Гос. номер", d["car_number"])
    d["car_mileage"] = c6.text_input("Пробег", d["car_mileage"])
    
    c7, c8 = st.columns(2)
    d["car_pts"] = c7.text_input("ПТС", d["car_pts"])
    d["car_sts"] = c8.text_input("СТС", d["car_sts"])
    
    c9, c10, c11 = st.columns(3)
    d["car_engine"] = c9.text_input("№ Двигателя", d["car_engine"])
    d["car_body"] = c10.text_input("№ Кузова", d["car_body"])
    d["car_frame"] = c11.text_input("№ Рамы (Шасси)", d["car_frame"])
    d["car_defects"] = st.text_input("Дефекты (недостатки)", d["car_defects"])

# --- ГЕНЕРАТОР WORD ДОКУМЕНТА ---
st.divider()

if st.button("📄 СГЕНЕРИРОВАТЬ ДОГОВОР И АКТ (.docx)", type="primary", use_container_width=True):
    doc = Document()
    
    # Настройки стиля
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(10)
    
    # Заголовок ДКП
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"ДОГОВОР КУПЛИ-ПРОДАЖИ АВТОМОБИЛЯ № {d['contract_num']}")
    r.bold = True
    
    # Город и дата
    p = doc.add_paragraph()
    p.add_run(f"{d['city']}").bold = False
    p.add_run(f"\t\t\t\t\t\t\t\t\t\t{d['contract_date']}")
    
    # Преамбула
    doc.add_paragraph(
        f"{d['buyer_company']}, в лице {d['buyer_rep']}, действующего(ей) на основании {d['buyer_poa']}, "
        f"именуемый(ая) в дальнейшем «Покупатель», с одной стороны, и, {d['seller_fio']}, именуемый(ая) "
        f"в дальнейшем «Продавец», с другой стороны, совместно именуемые «Стороны» заключили настоящий договор о нижеследующем:"
    )
    
    # 1. Предмет
    p = doc.add_paragraph()
    p.add_run("1. Предмет Договора").bold = True
    doc.add_paragraph(
        "1.1. Продавец обязуется передать в собственность Покупателя, а Покупатель обязуется принять и оплатить следующий "
        "бывший в эксплуатации автомобиль (далее по тексту - «Автомобиль»):\n"
        f"Марка (Модель): {d['car_mark']};\n"
        f"VIN: {d['car_vin']};\n"
        f"Модель, № двигателя: {d['car_engine']};\n"
        f"Номер кузова: {d['car_body']};\n"
        f"Номер рамы (шасси): {d['car_frame']};\n"
        f"Цвет: {d['car_color']};\n"
        f"Пробег (по показаниям одометра): {d['car_mileage']};\n"
        f"ПТС: {d['car_pts']};\n"
        f"Свидетельство о регистрации: {d['car_sts']};\n"
        f"Государственный регистрационный номер: {d['car_number']};\n"
        f"Год выпуска: {d['car_year']};\n"
        f"Дефекты (недостатки): {d['car_defects']}."
    )
    
    doc.add_paragraph(
        "1.2. Продавец гарантирует, что Автомобиль принадлежит ему на праве индивидуальной собственности, не заложен, не арендован, "
        "не является предметом спора, не состоит под арестом, не числится в розыске и не обременен никакими обязательствами перед третьими лицами."
    )
    doc.add_paragraph("1.3. Продавец гарантирует, что при заключении настоящего Договора уведомил Покупателя обо всех известных ему дефектах.")
    doc.add_paragraph("1.4. Продавец гарантирует, что на Автомобиле оригинальный пробег, и обслуживание осуществлялось надлежащим образом.")
    
    # 2. Цена
    p = doc.add_paragraph()
    p.add_run("2. Цена товара и порядок оплаты").bold = True
    doc.add_paragraph(f"2.1. Общая сумма Договора (стоимость Автомобиля) составляет {d['price_num']} рублей ({d['price_text']}), НДС не предусмотрен.")
    doc.add_paragraph(f"2.2. Расчет по договору осуществляется следующим образом:\n{d['payment_details']}.")
    
    # 3. Передача
    p = doc.add_paragraph()
    p.add_run("3. Передача Автомобиля").bold = True
    doc.add_paragraph("3.1. Автомобиль передается Покупателю в день подписания настоящего Договора.")
    doc.add_paragraph("3.2. Передача Автомобиля Покупателю оформляется актом приема-передачи, подписываемым Сторонами.")
    doc.add_paragraph("3.3. При приемке Автомобиля Покупатель осуществляет его проверку и сообщает Продавцу о недостатках.")
    doc.add_paragraph("3.4. Продавец передает Покупателю комплект документов, относящихся к Автомобилю.")
    doc.add_paragraph("3.5. Право собственности на Автомобиль переходит от Продавца к Покупателю с момента подписания Акта.")
    
    # 4. Прочие условия
    p = doc.add_paragraph()
    p.add_run("4. Прочие условия").bold = True
    doc.add_paragraph("4.1. Обстоятельства непреодолимой силы освобождают Стороны от обязательств на время действия таких обстоятельств.")
    doc.add_paragraph("4.2. Настоящий Договор вступает в силу с момента его подписания.")
    doc.add_paragraph("4.3. Настоящий Договор составлен в трех экземплярах, два экземпляра — для Покупателя, один — для Продавца.")

    # 6. Реквизиты
    p = doc.add_paragraph()
    p.add_run("6. Юридические адреса и банковские реквизиты Сторон").bold = True
    
    table = doc.add_table(rows=1, cols=2)
    table.autofit = True
    row = table.rows[0]
    
    # Колонка Продавца
    seller_cell = row.cells[0]
    seller_cell.text = f"ПРОДАВЕЦ:\n{d['seller_fio']}\nПаспорт: {d['seller_passport']}\nАдрес: {d['seller_address']}\nТел. {d['seller_phone']}\n\n________________________/____________/"
    
    # Колонка Покупателя
    buyer_cell = row.cells[1]
    buyer_cell.text = f"ПОКУПАТЕЛЬ:\n{d['buyer_company']}\nАдрес: {d['buyer_address']}\nТелефон: {d['buyer_phone']}\nИНН {d['buyer_inn']} КПП {d['buyer_kpp']}\nОГРН {d['buyer_ogrn']}\nБанк: {d['buyer_bank']}\nР/С {d['buyer_rs']}\nК/С {d['buyer_ks']}\nБИК {d['buyer_bik']}\n\n________________________/____________/"
    
    doc.add_page_break()
    
    # --- АКТ ПРИЕМА ПЕРЕДАЧИ ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("АКТ ПРИЕМА–ПЕРЕДАЧИ")
    r.bold = True
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"К ДОГОВОРУ КУПЛИ-ПРОДАЖИ № {d['contract_num']} от {d['contract_date']}")
    
    p = doc.add_paragraph()
    p.add_run(f"{d['city']}").bold = False
    p.add_run(f"\t\t\t\t\t\t\t\t\t\t{d['contract_date']}")
    
    doc.add_paragraph(
        f"{d['buyer_company']} в лице {d['buyer_rep']} действующего(ей) на основании {d['buyer_poa']}, "
        f"именуемый(ая) в дальнейшем «Покупатель», с одной стороны, и {d['seller_fio']}, именуемый(ая) "
        f"в дальнейшем «Продавец», с другой стороны, совместно именуемые «Стороны» заключили настоящий Акт о нижеследующем:"
    )
    
    doc.add_paragraph(
        f"1. Действуя в рамках Договора купли-продажи автомобиля № {d['contract_num']} от {d['contract_date']}, "
        f"заключенного между сторонами, Продавец передал, а Покупатель принял следующий товар:\n\n"
        f"Марка (Модель): {d['car_mark']};\n"
        f"VIN: {d['car_vin']};\n"
        f"Модель, № двигателя: {d['car_engine']};\n"
        f"Номер кузова: {d['car_body']};\n"
        f"Номер рамы (шасси): {d['car_frame']};\n"
        f"Цвет: {d['car_color']};\n"
        f"Пробег: {d['car_mileage']};\n"
        f"ПТС: {d['car_pts']};\n"
        f"Свидетельство о регистрации: {d['car_sts']};\n"
        f"Государственный регистрационный номер: {d['car_number']};\n"
        f"Год выпуска: {d['car_year']};\n"
        f"Дефекты (недостатки): {d['car_defects']}.\n"
    )
    
    doc.add_paragraph(f"2. Стоимость товара составляет {d['price_num']} рублей ({d['price_text']}), НДС не предусмотрен.")
    doc.add_paragraph("3. Вместе с товаром Продавец передал Покупателю:\n- Паспорт транспортного средства (ПТС)\n- Свидетельство регистрации ТC\n- Ключи от а/м\n")
    
    table_act = doc.add_table(rows=1, cols=2)
    row_act = table_act.rows[0]
    row_act.cells[0].text = f"Продавец\n\n_______________________/ {d['seller_fio']} /"
    row_act.cells[1].text = f"Покупатель\n\n______________________/ Представитель /"
    
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    
    st.download_button(
        label="💾 СКАЧАТЬ ГОТОВЫЙ ФАЙЛ (.docx)",
        data=target,
        file_name=f"ДКП_Акт_{d['car_mark']}_{d['car_vin']}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
