import streamlit as st
import pandas as pd
from datetime import datetime
import os

# Настройка страницы
st.set_page_config(page_title="Учет заказов: Пошив штор", page_icon="🪡", layout="wide")

PRICE_FILE = "tarifs.xlsx"
MATERIALS_FILE = "materials.xlsx"
PRODUCTS_FILE = "products.xlsx"

# 1. Автоматический справочник цен на РАБОТЫ (ИСПРАВЛЕНО: Полные, безопасные данные швейного цеха)
def init_default_price():
    default_data = {
        "Наименование работы": [
            "Пошив тюля с утяжелителем на шторной ленте", "Стянуть шторную ленту под размер",
            "Пошив портьеры на подкладке", "Пошив портьеры на подкладке из бархата", 
            "Надставка сверху и снизу", "Припуск на подгиб краев",
            "Пошив римской шторы без подкладки на тесьме", "Изготовление и притачивание канта", 
            "Евроуголок на кантах", "Чехол на молнии из бархата (до 50х50 см)", 
            "Изготовление и притачивание рулика"
        ],
        "Ед.изм": ["м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "кв.м.", "м.п.", "шт.", "шт.", "м.п."],
        "Цена (руб)": [420, 100, 770, 1078, 250, 150, 1450, 210, 350, 550, 180],
        "Наценка (%)": [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        "Категория": [
            "Тюль", "Тюль", 
            "Портьеры на подкладке", "Портьеры на подкладке", "Портьеры на подкладке", "Портьеры на подкладке",
            "Римская штора", "Римская штора", "Римская штора", 
            "Декоративная подушка", "Декоративная подушка"
        ]
    }
    df = pd.DataFrame(default_data)
    df.to_excel(PRICE_FILE, index=False)
    return df

# 2. Автоматический справочник цен на МАТЕРИАЛЫ И ФУРНИТУРУ (ИСПРАВЛЕНО: Безопасные данные)
def init_default_materials():
    default_data = {
        "Наименование материала": [
            "Шторная лента с кордами (8 см, 1:2)", "Утяжелитель нижний (стандарт)",
            "Ткань подкладки (стандарт)", "Шторная лента для портьер люкс",
            "Тесьма для римских штор", "Липучка велкро (комплект)", 
            "Кольца прозрачные (навеска)", "Внутренняя подушка (наполнитель)", 
            "Молния потайная (фурнитура)"
        ],
        "Ед.изм": ["м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "м.п.", "шт.", "шт.", "шт."],
        "Цена (руб)": [150, 250, 350, 200, 80, 120, 15, 300, 60],
        "Категория": [
            "Тюль", "Тюль", 
            "Портьеры на подкладке", "Портьеры на подкладке", 
            "Римская штора", "Римская штора", "Римская штора", 
            "Декоративная подушка", "Декоративная подушка"
        ]
    }
    df = pd.DataFrame(default_data)
    df.to_excel(MATERIALS_FILE, index=False)
    return df

# 3. Справочник НАИМЕНОВАНИЙ ИЗДЕЛИЙ
def init_default_products():
    default_data = {"Наименование изделия": ["Декоративная подушка", "Портьеры на подкладке", "Римская штора", "Тюль"]}
    df = pd.DataFrame(default_data)
    df.to_excel(PRODUCTS_FILE, index=False)
    return df

# Автоматическая проверка и генерация при первом запуске
if not os.path.exists(PRICE_FILE): df_price = init_default_price()
else:
    try: df_price = pd.read_excel(PRICE_FILE)
    except: df_price = init_default_price()

if not os.path.exists(MATERIALS_FILE): df_materials = init_default_materials()
else:
    try: df_materials = pd.read_excel(MATERIALS_FILE)
    except: df_materials = init_default_materials()

if not os.path.exists(PRODUCTS_FILE): df_products = init_default_products()
else:
    try: df_products = pd.read_excel(PRODUCTS_FILE)
    except: df_products = init_default_products()

PRODUCTS_LIST = sorted(df_products["Наименование изделия"].dropna().astype(str).tolist())
if "current_items" not in st.session_state: st.session_state.current_items = []
# =========================================================
# 🗂️ ЛЕВАЯ ПАНЕЛЬ (САЙДБАР)
# =========================================================
with st.sidebar:
    st.markdown("<h2 style='margin: 0;'>ЗАКАЗ №</h2>", unsafe_allow_html=True)
    num_col1, num_col2 = st.columns(2)
    with num_col1:
        order_digits = st.number_input("Номер", min_value=1, max_value=10000, value=101, step=1, label_visibility="collapsed", key="order_num_input")
    with num_col2:
        st.text_input("Год", value="- 2026", disabled=True, label_visibility="collapsed", key="year_input")
        
    st.markdown("---")
    st.markdown("**Дата**")
    order_date = st.date_input("Выбор даты", value=datetime.now().date(), key="date_input")
    project_name = st.text_input("Проект", placeholder="Например: Проект Роксана", key="project_input")
    client_name = st.text_input("ФИО заказчика", placeholder="ФИО заказчика", key="client_input")
    client_phone = st.text_input("Телефон", placeholder="+7...", key="phone_input")
    st.markdown("---")
    designer_name = st.text_input("Дизайнер проекта", placeholder="Укажите дизайнера", key="designer_input")
    admin_name = st.text_input("Менеджер/Админ", placeholder="Кто принял заказ", key="admin_input")

# =========================================================
# 🖥️ ОСНОВНОЕ ОКНО
# =========================================================
display_project_title = f" — «{project_name}»" if project_name else ""
st.markdown(f"## 🛠️ Расчет заказа № {order_digits}-2026{display_project_title}")

st.subheader("➕ Добавление позиции в расчет")
with st.container(border=True):
    col_group1, col_group2, col_group3 = st.columns(3)
    
    with col_group1:
        product_options = PRODUCTS_LIST + ["[ + Создать новое изделие ]"]
        selected_product = st.selectbox("📌 Наименование изделия:", product_options, index=0, key="prod_select")
        
        if selected_product == "[ + Создать новое изделие ]":
            item_group = st.text_input("Введите название изделия и нажмите Enter:", key="new_prod_text_direct").strip()
            if item_group and item_group not in PRODUCTS_LIST:
                new_p_df = pd.DataFrame([{"Наименование изделия": item_group}])
                df_products = pd.concat([df_products, new_p_df], ignore_index=True)
                df_products.to_excel(PRODUCTS_FILE, index=False)
                st.toast(f"Изделие '{item_group}' зафиксировано!", icon="💾")
                st.rerun()
            if item_group:
                st.markdown(f"Выбрано изделие: **`{item_group}`**")
        else:
            item_group = selected_product
            st.markdown(f"Выбрано изделие: **`{item_group}`**")

    with col_group2:
        entry_type = st.radio("⚡ Тип позиции:", ["🧵 Работа", "📦 Материал"], horizontal=True)

    with col_group3:
        if entry_type == "🧵 Работа":
            df_filtered = df_price[df_price["Категория"] == item_group] if "Категория" in df_price.columns else pd.DataFrame()
            sorted_works = sorted(df_filtered["Наименование работы"].dropna().astype(str).tolist()) if not df_filtered.empty else []
            filtered_dict = df_filtered.set_index("Наименование работы").to_dict(orient="index") if not df_filtered.empty else {}
            
            select_ops = sorted_works + ["[ + Добавить новое наименование работы ]"]
            selected_item = st.selectbox("📌 Наименование работы:", select_ops, key="work_select")
            is_new_entry = selected_item == "[ + Добавить новое наименование работы ]"
        else:
            df_filtered = df_materials[df_materials["Категория"] == item_group] if "Категория" in df_materials.columns else pd.DataFrame()
            filtered_dict = df_filtered.set_index("Наименование материала").to_dict(orient="index") if not df_filtered.empty else {}
            sorted_mats = sorted(list(filtered_dict.keys()))
            
            select_ops = sorted_mats + ["[ + Добавить новое наименование материала ]"]
            selected_item = st.selectbox("📌 Наименование материала/фурнитуры:", select_ops, key="mat_select")
            is_new_entry = selected_item == "[ + Добавить новое наименование материала ]"

    if is_new_entry:
        st.markdown("---")
        new_pos_name = st.text_input("Введите название НОВОЙ операции/материала и нажмите Enter:", key="new_pos_name_direct").strip()
        pos_name = new_pos_name
        if pos_name:
            st.markdown(f"Выбрана позиция: **`{pos_name}`**")
    else:
        pos_name = selected_item
        if pos_name:
            st.markdown(f"Выбрана позиция: **`{pos_name}`**")

    st.markdown("---")
    st.markdown("### 📏 Параметры и стоимость позиции")
    col_fields1, col_fields2, col_fields3, col_fields4 = st.columns(4)
    
    with col_fields1:
        st.text_input("Название позиции в расчете:", value=pos_name, disabled=True, key="pos_name_disabled")

    with col_fields2:
        unit_options = ["м.п.", "кв.м.", "шт."]
        if is_new_entry:
            unit_type = st.selectbox("Ед. измерения (Новое):", unit_options, key="unit_input_active")
        else:
            base_unit = filtered_dict[selected_item]["Ед.изм"] if selected_item in filtered_dict else "шт."
            st.text_input("Ед. измерения:", value=base_unit, disabled=True, key="unit_input_disabled")
            unit_type = base_unit

    with col_fields3:
        if is_new_entry:
            base_price = st.number_input("Цена руб. (Новая):", min_value=0, value=0, step=10, key="price_input_active")
        else:
            excel_price = int(filtered_dict[selected_item]["Цена (руб)"]) if selected_item in filtered_dict else 0
            st.number_input("Цена руб.:", value=excel_price, disabled=True, key="price_input_disabled")
            base_price = excel_price

    with col_fields4:
        qty = st.number_input("Количество:", min_value=0.01, value=1.0, step=0.1, key="qty_fields_input")

    col_calc1, col_calc2 = st.columns(2)
    with col_calc1:
        if entry_type == "🧵 Работа" and not is_new_entry:
            excel_markup = int(filtered_dict[selected_item].get("Наценка (%)", 0)) if selected_item in filtered_dict else 0
            markup = st.number_input("Наценка за сложность (%):", min_value=0, max_value=200, value=excel_markup, step=5, key="markup_calc")
        elif entry_type == "🧵 Работа" and is_new_entry:
            markup = st.number_input("Наценка за сложность (%):", min_value=0, max_value=200, value=0, step=5, key="markup_calc_new")
        else:
            markup = 0

    final_unit_price = round(base_price * (1 + markup / 100), 2)
    row_total = round(qty * final_unit_price, 2)
    
    with col_calc2:
        st.markdown(f"<h3 style='margin:0; padding-top:15px; color:#1f77b4;'>Стоимость позиции: {row_total:,} руб.</h3>", unsafe_allow_html=True)

    if st.button("📥 Добавить эту строчку в смету заказа", use_container_width=True, key="add_row_final_btn"):
        if is_new_entry and not pos_name:
            st.error("❌ Заполните название новой операции/материала!")
        elif not item_group:
            st.error("❌ Сначала выберите или введите изделие!")
        else:
            if is_new_entry:
                if entry_type == "🧵 Работа":
                    new_r = pd.DataFrame([{"Наименование работы": pos_name, "Ед.изм": unit_type, "Цена (руб)": base_price, "Наценка (%)": markup, "Категория": item_group}])
                    pd.concat([df_price, new_r], ignore_index=True).to_excel(PRICE_FILE, index=False)
                else:
                    new_r = pd.DataFrame([{"Наименование материала": pos_name, "Ед.изм": unit_type, "Цена (руб)": base_price, "Категория": item_group}])
                    pd.concat([df_materials, new_r], ignore_index=True).to_excel(MATERIALS_FILE, index=False)
            
            st.session_state.current_items.append({
                "Изделие": item_group,
                "Тип": "Работа" if entry_type == "🧵 Работа" else "Материал",
                "Описание позиции": pos_name,
                "Количество": qty,
                "Ед.изм": unit_type,
                "Цена, р": final_unit_price,
                "Стоимость, р": row_total
            })
            st.toast("Добавлено!", icon="✅")
            st.rerun()

# =========================================================
# 📊 ТАБЛИЦА СМЕТЫ С АВТОМАТИЧЕСКОЙ ГРУППИРОВКОЙ
# =========================================================
st.markdown("---")
st.subheader("📋 Состав и детализация текущего заказа")

if len(st.session_state.current_items) > 0:
    df_items = pd.DataFrame(st.session_state.current_items)
    unique_products = df_items["Изделие"].unique()
    
    for product in unique_products:
        st.markdown(f"#### 📦 Изделие: {product}")
        df_prod = df_items[df_items["Изделие"] == product].drop(columns=["Изделие"])
        st.table(df_prod)
        st.markdown(f"<p style='text-align: right; color: gray;'>Итого по изделию «{product}»: <b>{df_prod['Стоимость, р'].sum():,} р.</b></p>", unsafe_allow_html=True)
    
    st.divider()
    total_sum = df_items["Стоимость, р"].sum()
    st.success(f"### 💰 ВСЕГО К ОПЛАТЕ ПО ВСЕМ ИЗДЕЛИЯМ: {total_sum:,} руб.")
    if st.button("🗑️ Очистить всю смету", use_container_width=True, key="clear_all_btn"):
        st.session_state.current_items = []
        st.rerun()
else:
    st.info("В заказе пока пусто. Сформируйте параметры позиции выше и нажмите кнопку добавления.")
