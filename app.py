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

st.title("🚗 Генератор ДКП и Актов (по шаблонам Авто-К)")
st.caption("Финальная версия с исправленным скачиванием и точными шаблонами документов")

# --- БОКОВАЯ ПАНЕЛЬ: ЗАГРУЗКА И OCR ---
st.sidebar.header("📁 Загрузка документов")

uploaded_files = st.sidebar.file_uploader(
    "Фото (Паспорт, СТС, ПТС)",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True
)

def prepare_image(file_bytes):
    """Сжимаем фото, чтобы бесплатный OCR не выдавал таймаут"""
    img = Image.open(io.BytesIO(file_bytes))
    img.thumbnail((1500, 1500))
    buffer = io.BytesIO()
    img.convert("RGB").save(buffer, format="JPEG", quality=75)
    return buffer.getvalue()

# --- ИНИЦИАЛИЗАЦИЯ ДАННЫХ ---
today_str = datetime.datetime.now().strftime("%d.%m.%Y")

if "form_data" not in st.session_state:
    st.session_state.form_data = {
        "contract_num": "АК00000000",
        "contract_date": f"{today_str} г.",
        "city": "г. Калининград",
        "buyer_company": "ООО «Авто-К»",
        "buyer_rep": "Менеджера по закупкам автомобилей с пробегом Маркова Андрея Дмитриевича",
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
        "seller_fio": "",
        "seller_passport": "",
        "seller_address": "",
        "seller_phone": "",
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
        "price_num": "1 801 000",
        "price_text": "Один миллион восемьсот одна тысяча рублей 00 копеек",
        "payment_details": "денежные средства перечислить в качестве частичной оплаты приобретаемого Продавцом нового автомобиля..."
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
                
                vin_match = re.search(r'\b[A-HJ-NPR-Z0-9]{17}\b', full_text_upper)
                if vin_match:
                    st.session_state.form_data["car_vin"] = vin_match.group(0)
                    st.session_state.form_data["car_body"] = vin_match.group(0)

                number_match = re.search(r'[А-ЯA-Z]\s*\d{3}\s*[А-ЯA-Z]{2}\s*\d{2,3}', full_text_upper)
                if number_match:
                    st.session_state.form_data["car_number"] = number_match.group(0).replace(" ", "")

                year_match = re.search(r'\b(19|20)\d{2}\b', full_text)
                if year_match:
                    st.session_state.form_data["car_year"] = year_match.group(0)

                passport_match = re.search(r'\b\d{2}\s*\d{2}\s*\d{6}\b', full_text)
                if passport_match:
                    st.session_state.form_data["seller_passport"] = passport_match.group(0)

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
    d["payment_details"] = st.text_area("Особые условия оплаты (п. 2.2)", d["payment_details"], height=100)

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

# --- ФУНКЦИЯ ГЕНЕРАЦИИ WORD ---
def create_word_doc(data):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(10)
    
    # Заголовок ДКП
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"ДОГОВОР КУПЛИ-ПРОДАЖИ АВТОМОБИЛЯ № {data['contract_num']}").bold = True
    
    p = doc.add_paragraph()
    p.add_run(f"{data['city']}").bold = False
    p.add_run(f"\t\t\t\t\t\t\t\t\t\t\t\t{data['contract_date']}")
    
    doc.add_paragraph(
        f"{data['buyer_company']}, в лице {data['buyer_rep']}, действующего(ей) на основании {data['buyer_poa']}, "
        f"именуемый(ая) в дальнейшем «Покупатель», с одной стороны, и {data['seller_fio']}, именуемый(ая) "
        f"в дальнейшем «Продавец», с другой стороны, совместно именуемые «Стороны» заключили настоящий договор о нижеследующем:"
    )
    
    p = doc.add_paragraph()
    p.add_run("1. Предмет Договора").bold = True
    doc.add_paragraph(
        "1.1. Продавец обязуется передать в собственность Покупателя, а Покупатель обязуется принять и оплатить следующий "
        "бывший в эксплуатации автомобиль (далее по тексту - «Автомобиль»):\n"
        f"Марка (Модель): {data['car_mark']};\n"
        f"VIN: {data['car_vin']};\n"
        f"Модель, № двигателя: {data['car_engine']};\n"
        f"Номер кузова: {data['car_body']};\n"
        f"Номер рамы (шасси): {data['car_frame']};\n"
        f"Цвет: {data['car_color']};\n"
        f"Пробег (по показаниям одометра): {data['car_mileage']};\n"
        f"ПТС: {data['car_pts']};\n"
        f"Свидетельство о регистрации: {data['car_sts']};\n"
        f"Государственный регистрационный номер: {data['car_number']};\n"
        f"Год выпуска: {data['car_year']};\n"
        f"Дефекты (недостатки): {data['car_defects']}."
    )
    
    doc.add_paragraph(
        "1.2. Продавец гарантирует, что Автомобиль принадлежит ему на праве индивидуальной собственности, не заложен, не арендован, "
        "не является предметом спора, не состоит под арестом, не числится в розыске и не обременен никакими обязательствами перед третьими лицами."
    )
    doc.add_paragraph("1.3. Продавец гарантирует, что при заключении настоящего Договора уведомил Покупателя обо всех известных ему дефектах Автомобиля.")
    doc.add_paragraph("1.4. Продавец гарантирует, что на Автомобиле оригинальный пробег, и все техническое обслуживание осуществлялось надлежащим образом.")
    doc.add_paragraph("1.5. В случае нарушения Продавцом гарантий, Покупатель вправе в одностороннем порядке расторгнуть настоящий Договор.")
    
    p = doc.add_paragraph()
    p.add_run("2. Цена товара и порядок оплаты").bold = True
    doc.add_paragraph(f"2.1. Общая сумма Договора (стоимость Автомобиля) составляет {data['price_num']} рублей ({data['price_text']}), НДС не предусмотрен.")
    doc.add_paragraph(f"2.2. Расчет по договору осуществляется следующим образом:\n- {data['payment_details']}.")
    
    p = doc.add_paragraph()
    p.add_run("3. Передача Автомобиля").bold = True
    doc.add_paragraph("3.1. Автомобиль передается Покупателю в день подписания настоящего Договора.")
    doc.add_paragraph("3.2. Передача Автомобиля Покупателю оформляется актом приема-передачи, подписываемым Сторонами.")
    doc.add_paragraph("3.3. При приемке Автомобиля Покупатель за счет Продавца осуществляет его проверку.")
    doc.add_paragraph("3.4. Продавец передает Покупателю комплект документов, относящихся к Автомобилю.")
    doc.add_paragraph("3.5. Право собственности на Автомобиль переходит от Продавца к Покупателю с момента подписания Акта приема-передачи.")
    doc.add_paragraph("3.6. Продавец уведомлен о том, что Покупатель не регистрирует на свое имя транспортные средства, предназначенные для продажи.")
    
    p = doc.add_paragraph()
    p.add_run("4. Прочие условия").bold = True
    doc.add_paragraph("4.1. Обстоятельства непреодолимой силы освобождают Стороны от обязательств на время действия таких обстоятельств.")
    doc.add_paragraph("4.2. В случае, если выяснится, что Автомобиль является предметом спора, Продавец обязуется компенсировать убытки Покупателя.")
    doc.add_paragraph("4.3. Все споры будут по возможности решаться путем переговоров.")
    doc.add_paragraph("4.4. В случае невозможности разрешения разногласий, спор подлежит рассмотрению в суде.")
    doc.add_paragraph("4.5. Настоящий Договор вступает в силу с момента его подписания.")
    doc.add_paragraph("4.6. Настоящий Договор составлен в трех экземплярах, два экземпляра — для Покупателя, один — для Продавца.")

    p = doc.add_paragraph()
    p.add_run("6. Юридические адреса и банковские реквизиты Сторон").bold = True
    
    table = doc.add_table(rows=1, cols=2)
    row = table.rows[0]
    row.cells[0].text = f"ПРОДАВЕЦ:\n{data['seller_fio']}\nПаспорт: {data['seller_passport']}\nАдрес: {data['seller_address']}\nТел. {data['seller_phone']}\n\n________________________/____________/"
    row.cells[1].text = f"ПОКУПАТЕЛЬ:\n{data['buyer_company']}\nАдрес: {data['buyer_address']}\nТелефон: {data['buyer_phone']}\nИНН {data['buyer_inn']} КПП {data['buyer_kpp']}\nОГРН {data['buyer_ogrn']}\nБанк: {data['buyer_bank']}\nР/С {data['buyer_rs']}\nК/С {data['buyer_ks']}\nБИК {data['buyer_bik']}\n\n________________________/____________/"
    
    doc.add_paragraph("\nУважаемый Клиент! Благодарим Вас за выбор нашей компании. Обращаем Ваше внимание, что Вы вправе прекратить регистрацию Автомобиля в органах ГИБДД по своему заявлению по истечению 10 дней с момента перехода права собственности на Автомобиль.").italic = True
    
    doc.add_page_break()
    
    # --- АКТ ПРИЕМА ПЕРЕДАЧИ ---
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run("АКТ ПРИЕМА–ПЕРЕДАЧИ").bold = True
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run(f"К ДОГОВОРУ КУПЛИ-ПРОДАЖИ № {data['contract_num']} от {data['contract_date']}")
    
    p = doc.add_paragraph()
    p.add_run(f"{data['city']}").bold = False
    p.add_run(f"\t\t\t\t\t\t\t\t\t\t\t\t{data['contract_date']}")
    
    doc.add_paragraph(
        f"{data['buyer_company']} в лице {data['buyer_rep']} действующего(ей) на основании {data['buyer_poa']}, "
        f"именуемый(ая) в дальнейшем «Покупатель», с одной стороны, и {data['seller_fio']}, именуемый(ая) "
        f"в дальнейшем «Продавец», с другой стороны, совместно именуемые «Стороны» заключили настоящий Акт о нижеследующем:"
    )
    
    doc.add_paragraph(
        f"1. Действуя в рамках Договора купли-продажи автомобиля № {data['contract_num']} от {data['contract_date']}, "
        f"заключенного между сторонами, Продавец передал, а Покупатель принял следующий товар:\n\n"
        f"Марка (Модель): {data['car_mark']};\n"
        f"VIN: {data['car_vin']};\n"
        f"Модель, № двигателя: {data['car_engine']};\n"
        f"Номер кузова: {data['car_body']};\n"
        f"Номер рамы (шасси): {data['car_frame']};\n"
        f"Цвет: {data['car_color']};\n"
        f"Пробег: {data['car_mileage']};\n"
        f"ПТС: {data['car_pts']};\n"
        f"Свидетельство о регистрации: {data['car_sts']};\n"
        f"Государственный регистрационный номер: {data['car_number']};\n"
        f"Год выпуска: {data['car_year']};\n"
        f"Дефекты (недостатки): {data['car_defects']}.\n"
    )
    
    doc.add_paragraph(f"2. Стоимость товара составляет {data['price_num']} рублей ({data['price_text']}), НДС не предусмотрен.")
    doc.add_paragraph("3. Вместе с товаром Продавец передал Покупателю:\n- Паспорт транспортного средства (ПТС)\n- Свидетельство регистрации ТC\n- Ключи от а/м\n- Дополнительный комплект резины\n- Комплект ковриков в салон")
    
    table_act = doc.add_table(rows=1, cols=2)
    row_act = table_act.rows[0]
    row_act.cells[0].text = f"Продавец\n\n_______________________/ {data['seller_fio']} /"
    row_act.cells[1].text = f"Покупатель\n\n______________________/ Представитель /"
    
    target = io.BytesIO()
    doc.save(target)
    target.seek(0)
    return target

# --- КНОПКА ГЕНЕРАЦИИ И СКАЧИВАНИЯ ---
st.divider()

st.download_button(
    label="📄 СГЕНЕРИРОВАТЬ И СКАЧАТЬ ПОЛНЫЙ КОМПЛЕКТ (.docx)",
    data=create_word_doc(st.session_state.form_data),
    file_name=f"ДКП_Акт_{d['car_mark']}_{d['car_vin']}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    type="primary",
    use_container_width=True
)
